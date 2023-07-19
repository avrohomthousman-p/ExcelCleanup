using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;


namespace ExcelDataCleanup
{

    /// <summary>
    /// Implementation of the IMergeCleaner interface that seeks to maintain the exact same formatting after doing an
    /// unmerge. This is the primary implementation of this interface and should be the first tool used to clean merges.
    /// This implementation considers the table to start at the first row that contains at lease 3 non-empty cells.
    /// Cells are considered to be data cells if the column they are in contains text in the cell at the top of the table.
    /// All major and minor headers are not resized, but get set to not wrap their text, so excel can display it outside
    /// its cell boundry. Data cells are resized such that their size after the unmerge matches their original size before
    /// the unmerge.
    /// </summary>
    internal class PrimaryMergeCleaner : AbstractMergeCleaner
    {

        //number of non empty cells to be the first data row
        private static readonly int NUM_FULL_COLUMNS_REQUIRED = 3;


        private int firstRowOfTable;


        private bool[] isDataColumn;


        private Dictionary<int, double> originalColumnWidths = new Dictionary<int, double>();





        /// <inheritdoc/>
        protected override void FindTableBounds(ExcelWorksheet worksheet)
        {

            for (int row = 1; row <= worksheet.Dimension.Rows; row++)
            {
                if (IsDataRow(worksheet, row))
                {
                    firstRowOfTable = row;

                    Console.WriteLine("First data row is row " + row);

                    TrackDataColumns(worksheet);

                    return;
                }
            }


            //DEFAULT: if no data row is found there is something wrong
            throw new System.IO.InvalidDataException("Could not find data table in excel report.");
        }




        /// <summary>
        /// Checks if the specified row is a data row.
        /// </summary>
        /// <param name="worksheet">the worksheet where the row in question can be found</param>
        /// <param name="row">the row to be checked</param>
        /// <returns>true if the specified row is a data row and false otherwise</returns>
        private bool IsDataRow(ExcelWorksheet worksheet, int row)
        {

            int fullColumnsFound = 0;


            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                if (!IsEmptyCell(worksheet.Cells[row, col]))
                {
                    fullColumnsFound++;
                    if (fullColumnsFound == NUM_FULL_COLUMNS_REQUIRED)
                    {
                        return true;
                    }
                }
            }

            return false;
        }




        /// <summary>
        /// Populates the local variable isDataColumn with true's for each column that is a data column
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private void TrackDataColumns(ExcelWorksheet worksheet)
        {
            isDataColumn = new bool[worksheet.Dimension.End.Column];

            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {

                //a column is a data column if it has description text in the first row of the table.
                isDataColumn[col - 1] = !IsEmptyCell(worksheet.Cells[firstRowOfTable, col]);

            }
        }




        /// <inheritdoc/>
        protected override void UnMergeMergedSections(ExcelWorksheet worksheet)
        {

            RecordOriginalColumnWidths(worksheet);


            ExcelWorksheet.MergeCellsCollection mergedCells = worksheet.MergedCells;


            for (int i = mergedCells.Count - 1; i >= 0; i--)
            {
                var merged = mergedCells[i];


                //sometimes a change to one part of the worksheet causes a merge cell to stop
                //existing. The corrisponding entry in the merge collection to becomes null.
                if (merged == null)
                {
                    continue;
                }

                Console.WriteLine("merge at " + merged.ToString());

                UnMergeCells(worksheet, merged.ToString());
            }

        }



        /// <summary>
        /// Records the starting widths of all data columns that are merged in a dictionary for later use in resizing those columns
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        private void RecordOriginalColumnWidths(ExcelWorksheet worksheet)
        {

            for (int col = 1; col <= isDataColumn.Length; col++)
            {

                ExcelRange currentCell = worksheet.Cells[firstRowOfTable, col];


                if (!currentCell.Merge || IsEmptyCell(currentCell))
                {
                    continue;
                }


                currentCell = GetMergeCellByPosition(worksheet, firstRowOfTable, col);

                double width = GetWidthOfMergeCell(worksheet, currentCell);

                originalColumnWidths.Add(col, width);

                col += (CountMergeCellLength(currentCell) - 1);
            }
        }



        /// <summary>
        /// Finds the full ExcelRange object that contains the entire merge at the specified address. 
        /// In other words, the specified row and column point to a cell that is merged to be part of a 
        /// larger cell. This method returns the ExcelRange for the ENTIRE merge cell.
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        /// <param name="row">the row of a cell that is part of the larger merge</param>
        /// <param name="col">the column of a cell that is part of the larger merge</param>
        /// <returns>the Excel range object containing the entire merge</returns>
        private ExcelRange GetMergeCellByPosition(ExcelWorksheet worksheet, int row, int col)
        {
            int index = worksheet.GetMergeCellId(row, col);
            string cellAddress = worksheet.MergedCells[index - 1];
            return worksheet.Cells[cellAddress];
        }



        /// <summary>
        /// Calculates the width of a merge cell.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="currentCells">the merge cell being mesured</param>
        /// <returns>the width of the specified merged cell</returns>
        private double GetWidthOfMergeCell(ExcelWorksheet worksheet, ExcelRange currentCells)
        {
            double width = 0;

            for (int col = currentCells.Start.Column; col <= currentCells.End.Column; col++)
            {
                width += worksheet.Column(col).Width; //alt:  currentCell.EntireColumn.Width;  
            }


            return width;
        }



        /// <summary>
        /// Counts the number of cells in a merge cell. This method only counts the number of cells
        /// in the first row of the merge.
        /// </summary>
        /// <param name="mergeCell">the address of the full merge cell</param>
        /// <returns>the number of cells in the merge</returns>
        private int CountMergeCellLength(ExcelRange mergeCell)
        {
            return mergeCell.End.Column - mergeCell.Start.Column + 1;
        }




        /// <summary>
        /// Unmerges the specified segment of merged cells.
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        /// <param name="cellAddress">the address of the ENTIRE merged section (eg A18:F24)</param>
        /// <returns>true if the specified cell was unmerged, and false otherwise</returns>
        private bool UnMergeCells(ExcelWorksheet worksheet, string cellAddress)
        {

            ExcelRange currentCells = worksheet.Cells[cellAddress];


            //record the style we had before any changes were made
            ExcelStyle originalStyle = currentCells.Style;




            MergeType mergeType = GetCellMergeType(currentCells);


            switch (mergeType)
            {
                case MergeType.NOT_A_MERGE:
                    return false;

                case MergeType.EMPTY:
                    break;

                case MergeType.MAIN_HEADER:
                    currentCells.Style.WrapText = false;
                    currentCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ConvertContentsToText(currentCells); //Ensure that dates are displayed correctly
                    Console.WriteLine("major header at " + currentCells.Address);
                    break;

                case MergeType.MINOR_HEADER:
                    currentCells.Style.WrapText = false;
                    break;

                case MergeType.DATA:
                    break;
            }


            //turning off custom hieght before unmerge will allow row to resize itself to fit everything
            worksheet.Row(currentCells.Start.Row).CustomHeight = false;




            //unmerge range
            currentCells.Merge = false;




            //restore the original style
            SetCellStyles(currentCells, originalStyle);
            


            //If there is more than one line of text in a header, it should be split into
            //multiple seperate headers.
            if(mergeType == MergeType.MAIN_HEADER)
            {
                SplitHeaderIntoMultipleRows(worksheet, currentCells);
            }



            return true;

        }




        /// <summary>
        /// Gets the type of merge that is found in the specified cell
        /// </summary>
        /// <param name="cell">the cell whose merge type is being checked</param>
        /// <returns>the MergeType object that corrisponds to the type of merge cell we are given</returns>
        private MergeType GetCellMergeType(ExcelRange cell)
        {
            if (cell.Merge == false)
            {
                return MergeType.NOT_A_MERGE;
            }

            if (IsEmptyCell(cell))
            {
                return MergeType.EMPTY;
            }

            if (IsMajorHeader(cell))
            {
                return MergeType.MAIN_HEADER;
            }

            if (IsMinorHeader(cell))
            {
                return MergeType.MINOR_HEADER;
            }

            if (IsDataCell(cell))
            {
                return MergeType.DATA;
            }


            throw new ArgumentException("Cannot determan merge type of specified cell");
        }




        /// <inheritdoc/>
        protected override void ResizeCells(ExcelWorksheet worksheet)
        {
            //resize all columns
            foreach (KeyValuePair<int, double> entry in originalColumnWidths)
            {
                worksheet.Column(entry.Key).Width = entry.Value;
            }
        }




        /// <inheritdoc/>
        protected override void DeleteColumns(ExcelWorksheet worksheet)
        {

            //We don't want to delete any columns before the first data column becuase that might 
            //mess up the whitespace around minor headers
            int firstDataColumn = FindFirstDataColumn(worksheet);


            for (int col = worksheet.Dimension.Columns; col > firstDataColumn; col--)
            {
                if (SafeToDeleteColumn(worksheet, col))
                {
                    PrepareColumnForDeletion(worksheet, col);
                    worksheet.DeleteColumn(col);

                    Console.WriteLine("Column " + col + " is being deleted");
                }
            }
        }



        /// <summary>
        /// Finds the column number of the first column that is considered a data column
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <returns>the column number of the first data column</returns>
        private int FindFirstDataColumn(ExcelWorksheet worksheet)
        {
            for (int i = 0; i < isDataColumn.Length; i++)
            {
                if (isDataColumn[i])
                {
                    return i + 1;
                }
            }


            return 1; //Default: first column
        }



        /// <summary>
        /// Checks if a column is safe to delete becuase it is empty other than possibly having major headers in it.
        /// </summary>
        /// <param name="worksheet">the worksheet where the column can be found</param>
        /// <param name="col">the column being checked</param>
        /// <returns>true if it is safe to delete the column and false if deleting it would result in data loss</returns>
        private bool SafeToDeleteColumn(ExcelWorksheet worksheet, int col)
        {
            for (int row = firstRowOfTable; row <= worksheet.Dimension.Rows; row++)
            {
                if (!IsEmptyCell(worksheet.Cells[row, col]))
                {
                    return false;
                }
            }


            return true;
        }



        /// <summary>
        /// Moves all major headers in the specified column to the column adjacent on the left, or right if we are on the  
        /// first column.
        /// </summary>
        /// <param name="worksheet">the worksheet the column could be found in</param>
        /// <param name="col">the column number we are preparing to delete</param>
        private void PrepareColumnForDeletion(ExcelWorksheet worksheet, int col)
        {
            for (int row = 1; row < firstRowOfTable; row++)
            {
                if (!IsEmptyCell(worksheet.Cells[row, col]))
                {

                    int destinationColumn;
                    if (col == 1) //we are on the first column
                    {
                        destinationColumn = 2; //move header right
                    }
                    else
                    {
                        destinationColumn = col - 1; //move header left
                    }

                    ExcelRange originCell = worksheet.Cells[row, col];
                    ExcelRange destinationCell = worksheet.Cells[row, destinationColumn];


                    //copy all styles and formatting
                    originCell.CopyStyles(destinationCell);

                    //Move the text to the destination cell (store it as a string to avoid excel display issues with dates)
                    destinationCell.SetCellValue(0, 0, originCell.Text);


                    originCell.Value = null;
                }
            }
        }




        /// <inheritdoc/>
        protected override bool IsDataCell(ExcelRange cell)
        {

            return isDataColumn[cell.Start.Column - 1];

        }



        /// <inheritdoc/>
        protected override bool IsInsideTable(ExcelRange cell)
        {

            return cell.Start.Row >= firstRowOfTable;

        }



        /// <inheritdoc/>
        protected override bool IsMajorHeader(ExcelRange cell)
        {
            return !IsEmptyCell(cell) && cell.Start.Row < firstRowOfTable;
        }



        /// <inheritdoc/>
        protected override bool IsMinorHeader(ExcelRange cells)
        {
            if (IsEmptyCell(cells) || !IsInsideTable(cells))
            {
                return false;
            }



            return !isDataColumn[cells.Start.Column - 1];
        }
    }
}

