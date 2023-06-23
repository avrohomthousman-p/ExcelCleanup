using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ExcelDataCleanup
{
    public class FileCleaner
    {

        //Needed for determaning a cell width based on its text.
        private static readonly int DEFAULT_FONT_SIZE = 10;


        private static int firstRowOfTable;


        private static bool[] isDataColumn;




        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]




        public static void Main(string[] args)
        {
            string filepath = "";

            if (args != null && args.Count() > 0)
            {
                filepath = args[0];
            }

            else
            {

                // C:\Users\avroh\Downloads\ExcelProject\PayablesAccountReport_large.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\PayablesAccountReport_1Prop.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\ReportPayablesRegister.xlsx

                // C:\Users\avroh\Downloads\ExcelProject\ProfitAndLossStatementDrillthrough.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\AgedReceivables.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\LedgerExport.xlsx

                // C:\Users\avroh\Downloads\ExcelProject\TrialBalance.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\ProfitAndLossStatementByPeriod.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\testFile.xlsx


                Console.WriteLine("Please enter the filepath of the Excel report you want to clean:");
                filepath = Console.ReadLine();

                /*
                OpenFileDialog dialog = new OpenFileDialog();
                if (DialogResult.OK == dialog.ShowDialog())
                {
                    filepath = dialog.FileName;
                }
                Console.WriteLine("Hello World!");
                */
            }

            

            OpenXLSX( ConvertFileToBytes(filepath), filepath );
        }




        /// <summary>
        /// Opens the specified file and writes its contents to a byte array. This function is only needed for testing. In production
        /// the file itself will be passed in as a byte array, not as a filepath.
        /// </summary>
        /// <param name="filepath">the location of the file</param>
        /// <returns>a byte array with the contents of the file in it</returns>
        private static byte[] ConvertFileToBytes(string filepath)
        {
            FileInfo existingFile = new FileInfo(filepath);
            byte[] fileData = new byte[existingFile.Length];


            var fileStream = existingFile.Open(FileMode.Open);
            int bytesRead = 0;
            int bytesToRead = (int) existingFile.Length;
            while (bytesToRead > 0)
            {
                int justRead = fileStream.Read(fileData, bytesRead, bytesToRead);

                if(justRead == 0)
                {
                    break;
                }

                bytesRead += justRead;
                bytesToRead -= justRead;
            }


            return fileData;
        }




        /// <summary>
        /// Opens an existing excel file and reads some values and properties
        /// </summary>
        /// <param name="file">the excel file in byte form</param>
        /// <param name="originalFileName">the file name of the original excel file</param>
        public static void OpenXLSX(byte[] file, string originalFileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            using (ExcelPackage package = new ExcelPackage(new MemoryStream(file)))
            {
                //Get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];


                DeleteHiddenRows(worksheet);


                RemoveAllHyperLinks(worksheet);


                RemoveAllMerges(worksheet);


                FixExcelTypeWarnings(worksheet);





                package.SaveAs(originalFileName.Replace(".xlsx", "_fixed.xlsx"));

            }

            Console.WriteLine("Workbook Cleanup complete");
            Console.Read();
        }




        /// <summary>
        /// Deletes all hidden rows in the specified worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private static void DeleteHiddenRows(ExcelWorksheet worksheet)
        {
            var start = worksheet.Dimension.Start;
            var end = worksheet.Dimension.End;

            for (int row = end.Row; row >= start.Row; row--)
            {
                if (worksheet.Row(row).Hidden == true)
                {
                    worksheet.DeleteRow(row);
                    Console.WriteLine("Deleted Hidden Row : " + row);
                }
            }
        }



        /// <summary>
        /// Removes all hyperlinks that are in any of the cells in the specified worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private static void RemoveAllHyperLinks(ExcelWorksheet worksheet)
        {
            var start = worksheet.Dimension.Start;
            var end = worksheet.Dimension.End;

            for (int row = end.Row; row >= start.Row; row--)
            {
                for (int col = start.Column; col <= end.Column; ++col)
                {

                    var cell = worksheet.Cells[row, col];
                    StripCellOfHyperLink(cell, row, col);

                }
            }
        }




        /// <summary>
        /// Removes hyperlinks in the specified Excel Cell if any are present.
        /// </summary>
        /// <param name="cell">the cell whose hyperlinks should be removed</param>
        /// <param name="row">the row the cell is in</param>
        /// <param name="col">the column the cell is in</param>
        private static void StripCellOfHyperLink(ExcelRange cell, int row, int col)
        {
            if (cell.Hyperlink != null)
            {
                //worksheet.Cells[cell.EntireColumn.ToString()].Merge = false;
                //cell.Hyperlink.ReferenceAddress("");

                Console.WriteLine("Row=" + row.ToString() + " Col=" + col.ToString() + " Hyperlink=" + cell.Hyperlink);
                //  Uri uval = new Uri(cell.Text, UriKind.Relative);
                // cell.Hyperlink;
                var val = cell.Value;
                cell.Hyperlink = null;
                ////cell.Hyperlink = new Uri(cell.ToString(), UriKind.Absolute);
                cell.Value = val;
            }
        }




        /// <summary>
        /// Removes all merge cells from the specified worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private static void RemoveAllMerges(ExcelWorksheet worksheet)
        {

            FindTableBounds(worksheet);

            Dictionary<int, double> originalColumnWidths = RecordOriginalColumnWidths(worksheet);

            UnMergeMergedSections(worksheet);

            ResizeColumns(worksheet, originalColumnWidths);

            DeleteColumns(worksheet);
        }




        /// <summary>
        /// Finds the first row that is considered part of the table in the specified worksheet and saves the 
        /// row number to this classes local variable (topTableRow) for later use
        /// </summary>
        /// <param name="worksheet">the worksheet we are working on</param>
        private static void FindTableBounds(ExcelWorksheet worksheet)
        {

            for (int row = 1; row <= worksheet.Dimension.Rows; row++)
            {
                if(IsDataRow(worksheet, row))
                {
                    firstRowOfTable = row;

                    Console.WriteLine("First data row is row " + row);

                    TrackDataColumns(worksheet);

                    return;
                }
            }


            //DEFAULT: if no data row is found there is something wrong
            throw new Exception("Could not find data table in excel report.");
        }




        /// <summary>
        /// Checks if the specified row is a data row.
        /// 
        /// Current Implementation: a row is a data row if it contains at least 3 cells with text
        /// </summary>
        /// <param name="worksheet">the worksheet where the row in question can be found</param>
        /// <param name="row">the row to be checked</param>
        /// <returns>true if the specified row is a data row and false otherwise</returns>
        private static bool IsDataRow(ExcelWorksheet worksheet, int row)
        {

            const int NUM_FULL_COLUMNS_REQUIRED = 3;
            int fullColumnsFound = 0;


            for(int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                if (!IsEmptyCell(worksheet.Cells[row, col]))
                {
                    fullColumnsFound++;
                    if(fullColumnsFound == NUM_FULL_COLUMNS_REQUIRED)
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
        private static void TrackDataColumns(ExcelWorksheet worksheet)
        {
            isDataColumn = new bool[worksheet.Dimension.End.Column];

            for(int col = 1; col <= worksheet.Dimension.Columns; col++)
            {

                //a column is a data column if it has description text in the first row of the table.
                isDataColumn[col - 1] = !IsEmptyCell(worksheet.Cells[firstRowOfTable, col]);

            }
        }



        /// <summary>
        /// Records the starting widths of all data columns that are merged in a dictionary for later use in resizing those columns
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <returns>a dictionary containing all the column numbers (of data columns only) and their original widths 
        /// (before any unmerging was done)</returns>
        private static Dictionary<int, double> RecordOriginalColumnWidths(ExcelWorksheet worksheet)
        {
            Dictionary<int, double> columnWidths = new Dictionary<int, double>();


            for(int col = 1; col <= isDataColumn.Length; col++)
            {

                ExcelRange currentCell = worksheet.Cells[firstRowOfTable, col];


                if (!currentCell.Merge || IsEmptyCell(currentCell))
                {
                    continue;
                }


                currentCell = GetMergeCellByPosition(worksheet, firstRowOfTable, col);

                double width = GetWidthOfMergeCell(worksheet, currentCell);

                columnWidths.Add(col, width);

                col += (CountMergeCellLength(currentCell) - 1);
            }


            return columnWidths;
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
        private static ExcelRange GetMergeCellByPosition(ExcelWorksheet worksheet, int row, int col)
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
        private static double GetWidthOfMergeCell(ExcelWorksheet worksheet, ExcelRange currentCells)
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
        private static int CountMergeCellLength(ExcelRange mergeCell)
        {
            return mergeCell.End.Column - mergeCell.Start.Column + 1;
        }




        /// <summary>
        /// Unmerges all the merged sections in the worksheet.
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private static void UnMergeMergedSections(ExcelWorksheet worksheet)
        {

            ExcelWorksheet.MergeCellsCollection mergedCells = worksheet.MergedCells;


            for (int i = mergedCells.Count - 1; i >= 0; i--)
            {
                var merged = mergedCells[i];


                //sometimes a change to one part of the worksheet causes a merge cell to stop
                //existing. The corrisponding entry in the merge collection to becomes null.
                if(merged == null)
                {
                    continue;
                }

                Console.WriteLine("merge at " + merged.ToString());

                UnMergeCells(worksheet, merged.ToString());
            }

        }



        /// <summary>
        /// Unmerges the specified segment of merged cells.
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        /// <param name="cellAddress">the address of the ENTIRE merged section (eg A18:F24)</param>
        /// <returns>true if the specified cell was unmerged, and false otherwise</returns>
        private static bool UnMergeCells(ExcelWorksheet worksheet, string cellAddress)
        {

            ExcelRange currentCells = worksheet.Cells[cellAddress];




            //Sometimes unmerging a cell changes the row height. We need to reset it to its starting value
            double initialHeigth = worksheet.Row(currentCells.Start.Row).Height;

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
                    initialHeigth = GetHeightOfMergeCell(currentCells); //main headers sometimes span multiple rows
                    currentCells.Style.WrapText = false;
                    currentCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    Console.WriteLine("major header at " + currentCells.Address);
                    break;

                case MergeType.MINOR_HEADER:
                    currentCells.Style.WrapText = false;
                    break;

                case MergeType.DATA:
                    break;
            }     


            //unmerge range
            currentCells.Merge = false; 


            //restore original height
            worksheet.Row(currentCells.Start.Row).Height = initialHeigth;


            //restore the original style
            SetCellStyles(currentCells, originalStyle);


            return true;

        }




        /// <summary>
        /// Gets the type of merge that is found in the specified cell
        /// </summary>
        /// <param name="cell">the cell whose merge type is being checked</param>
        /// <returns>the MergeType object that corrisponds to the type of merge cell we are given</returns>
        private static MergeType GetCellMergeType(ExcelRange cell)
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





        /// <summary>
        /// Checks if the specified cell is a major header
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the specified cell contains a major header, and false otherwise</returns>
        private static bool IsMajorHeader(ExcelRange cell)
        {
            return !IsEmptyCell(cell) && cell.Start.Row < firstRowOfTable;
        }



        /// <summary>
        /// Checks if the specified cell is considered a minor header.
        /// 
        /// A minor header is defined as a merge cell that contains non-data text and is inside the table.
        /// </summary>
        /// <param name="cells">the cells that we are checking</param>
        /// <returns>true if the specified cells are a minor header and false otherwise</returns>
        private static bool IsMinorHeader(ExcelRange cells)
        {
            if (IsEmptyCell(cells) || !IsInsideTable(cells))
            {
                return false;
            }



            return !isDataColumn[cells.Start.Column - 1];
        }



        /// <summary>
        /// Checks if a cell has no text
        /// </summary>
        /// <param name="currentCells">the cell that is being checked for text</param>
        /// <returns>true if there is no text in the cell, and false otherwise</returns>
        private static bool IsEmptyCell(ExcelRange currentCells)
        {
            return currentCells.Text == null || currentCells.Text.Length == 0;
        }



        /// <summary>
        /// Checks if the cell at the specified coordinates is a data cell. This is used by the
        /// current implementation: a cell is a data row if it has a $ in it.
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell is a data cell and false otherwise</returns>
        private static bool IsDataCell(ExcelRange cell)
        {

            return isDataColumn[cell.Start.Column - 1];

        }



        /// <summary>
        /// Checks if the specified cell is inside the table in the worksheet, and not a header 
        /// above the table
        /// </summary>
        /// <param name="cell">the cell whose location is being checked</param>
        /// <returns>true if the specified cell is inside a table and false otherwise</returns>
        private static bool IsInsideTable(ExcelRange cell)
        {

            return cell.Start.Row >= firstRowOfTable;

        }




        /// <summary>
        /// Counts up the total height of all rows in the specifed merge cell
        /// </summary>
        /// <param name="currentCell">the merge cell whose height is being mesured</param>
        /// <returns>the height of the specified merge cell</returns>
        private static double GetHeightOfMergeCell(ExcelRange currentCell)
        {

            ExcelWorksheet worksheet = currentCell.Worksheet;

            double height = 0;
            for(int row = currentCell.Start.Row; row <= currentCell.End.Row; row++)
            {
                height += worksheet.Row(row).Height;
            }


            return height;
        }




        /// <summary>
        /// Sets the PatternType, Color, Border, Font, and Horizontal Alingment of all the cells
        /// in the specifed range.
        /// </summary>
        /// <param name="currentCells">the cells whose style must be set</param>
        /// <param name="style">all the styles we should use</param>
        private static void SetCellStyles(ExcelRange currentCells, ExcelStyle style)
        {


            //Ensure the color of the cells dont get changed
            currentCells.Style.Fill.PatternType = style.Fill.PatternType;
            if (currentCells.Style.Fill.PatternType != ExcelFillStyle.None)
            {
                currentCells.Style.Fill.PatternColor.SetColor(
                    GetColorFromARgb(style.Fill.PatternColor.LookupColor()));

                currentCells.Style.Fill.BackgroundColor.SetColor(
                    GetColorFromARgb(style.Fill.BackgroundColor.LookupColor()));
            }


            currentCells.Style.Border = style.Border;
            currentCells.Style.Font.Bold = style.Font.Bold;
            currentCells.Style.Font.Size = style.Font.Size;
            currentCells.Style.Font.Name = style.Font.Name;
            currentCells.Style.Font.Scheme = style.Font.Scheme;
            currentCells.Style.Font.Charset = style.Font.Charset;


            //Not sure why but it only sucsessfully sets these settings if these 2 lines are NOT executed
            //currentCells.Style.WrapText = style.WrapText;
            //currentCells.Style.HorizontalAlignment = style.HorizontalAlignment;
        }



        /// <summary>
        /// Generates A Color Object from an ARGB string.
        /// </summary>
        /// <param name="argb">the argb code of the color needed</param>
        /// <returns>an instance of System.Drawing.Color that matches the specified argb code</returns>
        private static System.Drawing.Color GetColorFromARgb(String argb)
        {
            if (argb.StartsWith("#"))
            {
                argb = argb.Substring(1);
            }
            else if (argb.StartsWith("0x"))
            {
                argb = argb.Substring(2);
            }


            System.Drawing.Color color = System.Drawing.Color.FromArgb(
                int.Parse(argb, System.Globalization.NumberStyles.HexNumber));


            return color;
        }



        /// <summary>
        /// Deletes all empty columns in the worksheet created by unmerges
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private static void DeleteColumns(ExcelWorksheet worksheet)
        {

            //We don't want to delete any columns before the first data column becuase that might 
            //mess up the whitespace around minor headers
            int firstDataColumn = FindFirstDataColumn(worksheet);


            for(int col = worksheet.Dimension.Columns; col > firstDataColumn;  col--)
            {
                if(SafeToDeleteColumn(worksheet, col))
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
        private static int FindFirstDataColumn(ExcelWorksheet worksheet)
        {
            for(int i = 0; i < isDataColumn.Length; i++)
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
        private static bool SafeToDeleteColumn(ExcelWorksheet worksheet, int col)
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
        private static void PrepareColumnForDeletion(ExcelWorksheet worksheet, int col)
        {
            for(int row = 1; row < firstRowOfTable; row++)
            {
                if (!IsEmptyCell(worksheet.Cells[row, col]))
                {
                    
                    int destinationColumn;
                    if(col == 1) //we are on the first column
                    {
                        destinationColumn = 2; //move header right
                    }
                    else
                    {
                        destinationColumn = col - 1; //move header left
                    }

                    ExcelRange originCell = worksheet.Cells[row, col];
                    ExcelRange destinationCell = worksheet.Cells[row, destinationColumn];


                    //Move the text and its formatting
                    destinationCell.Value = originCell.Value;
                    destinationCell.Style.Font.Name = originCell.Style.Font.Name;
                    destinationCell.Style.Font.Bold = originCell.Style.Font.Bold;
                    destinationCell.Style.Font.Size = originCell.Style.Font.Size;

                    originCell.Value = null;
                }
            }
        }




        /// <summary>
        /// Resizes the columns specified in the dictionary to the size stored in the dictionary
        /// </summary>
        /// <param name="worksheet">the worksheet curently being cleaned</param>
        /// <param name="widthsToUse">a dictionary mapping column numbers to desired widths</param>
        private static void ResizeColumns(ExcelWorksheet worksheet, Dictionary<int, double> widthsToUse)
        {
            foreach (KeyValuePair<int, double> entry in widthsToUse)
            {
                worksheet.Column(entry.Key).Width = entry.Value;
            }
        }




        /// <summary>
        /// Checks all cells in the worksheet for numbers that are being stored as text, and replaces them with the actual number.
        /// The purpose of this is to remove the excel warning that comes up when numbers are stored as text.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        private static void FixExcelTypeWarnings(ExcelWorksheet worksheet)
        {
            for (int i = 1; i <= worksheet.Dimension.Rows; i++)
            {
                for (int j = 1; j <= worksheet.Dimension.Columns; j++)
                {

                    ExcelRange cell = worksheet.Cells[i, j];

                    double? data = ConvertToNumber(cell.Text);

                    if (data != null)
                    {

                        cell.Value = data; //Replace the cell data with the same thing just not in text form


                        //When the alingment is set to general, text is left aligned but numbers are right aligned.
                        //Therefore if we change from text to number and we want to maintain alignment, we need to 
                        //change to right aligned.
                        if (cell.Style.HorizontalAlignment.Equals(ExcelHorizontalAlignment.General))
                        {
                            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                    }
                }
            }
        }



        /// <summary>
        /// Attempts to convert that specified string into a double
        /// </summary>
        /// <param name="data">the text that should be converted to a number</param>
        /// <returns>the text as a double object or null if it could not be converted</returns>
        private static double? ConvertToNumber(string data)
        {

            double result;

            bool sucsess = Double.TryParse(data, out result);


            return (sucsess ? result : null);
        }

    }

}
