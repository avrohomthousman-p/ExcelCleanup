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


        enum MergeType { NOT_A_MERGE, EMPTY, MAIN_HEADER, MINOR_HEADER, DATA }


        private static int topTableRow;


        //it is always allowed to increase a column's width by this amount or less, as it is considered 
        //minor and is unlikly to mess up any formatting
        private static readonly int MINOR_WIDTH_INCREASE = 3;


        //It is always allowed to set the width of a column to have space for this many characters.
        private static readonly int SMALL_TEXT_LENGTH = 6;


        //Some data that is needed for font size conversions:
        private static readonly double DEFAULT_FONT_SIZE = 10;


        //Stores column numbers of columns that are potentially safe to delete, as before the unmerge
        //they were part of data cells.
        private static HashSet<int> columnsToDelete = new HashSet<int>();



        //Dictionary to track all the columns that need to be resized and the size they should be.
        private static Dictionary<int, double> desiredColumnSizes = new Dictionary<int, double>();


        //Tracks which rows (might) need to be resized
        private static bool[] rowNeedsResize;




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

            OpenXLSX(filepath);
        }




        /// <summary>
        /// Opens an existing excel file and reads some values and properties
        /// </summary>
        public static void OpenXLSX(string Filepath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo existingFile = new FileInfo(Filepath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //Get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];



                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                for (int row = end.Row; row >= start.Row; row--)
                {
                    if (worksheet.Row(row).Hidden == true)
                    {
                        worksheet.DeleteRow(row);
                        Console.WriteLine("Deleted Hidden Row : " + row);
                        continue;
                    }



                    UnmergeEntireRow(worksheet.Row(row));





                    for (int col = start.Column; col <= end.Column; ++col)
                    {
                        var cell = worksheet.Cells[row, col];

                        if (cell.Value != null)
                            Console.WriteLine("Row=" + row.ToString() + " Col=" + col.ToString() + " Value=" + cell.Value);



                        RemoveHyperLink(cell, row, col);
                    }
                }



                FindTableBounds(worksheet);

                UnMergeMergedSections(worksheet);

                ResizeColumns(worksheet);

                ResizeRows(worksheet);


                foreach (int i in columnsToDelete)
                {
                    Console.WriteLine("column " + i + " was marked for deletion");
                }


                DeleteColumns(worksheet);

                FixExcelTypeWarnings(worksheet);





                package.SaveAs(Filepath.Replace(".xlsx", "_fixed.xlsx"));

            }

            Console.WriteLine("Workbook Cleanup complete");
            Console.Read();
        }



        /// <summary>
        /// Unmerges the entire specified row if it is already merged, otherwise does nothing
        /// </summary>
        /// <param name="row">the row to be unmerged</param>
        private static void UnmergeEntireRow(ExcelRow row)
        {
            if (row.Merged)
            {
                row.Merged = false;
            }
        }



        /// <summary>
        /// Removes hyperlinks in the specified Excel Cell if any are present.
        /// </summary>
        /// <param name="cell">the cell whose hyperlinks should be removed</param>
        /// <param name="row">the row the cell is in</param>
        /// <param name="col">the column the cell is in</param>
        private static void RemoveHyperLink(ExcelRange cell, int row, int col)
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
        /// Finds the first row of the table in the specified worksheet and saves the 
        /// row number to this classes local variable (topTableRow) for later use
        /// </summary>
        /// <param name="worksheet">the worksheet we are working on</param>
        private static void FindTableBounds(ExcelWorksheet worksheet)
        {
            for (int i = 1; i <= worksheet.Dimension.Rows; i++)
            {
                for (int j = 1; j <= worksheet.Dimension.Columns; j++)
                {
                    if (isDataCell(worksheet.Cells[i, j]))
                    {
                        j = FindRightSideOfTable(worksheet, i, j);
                        i = FindTopEdgeOfTable(worksheet, i, j);
                        topTableRow = i;
                        Console.WriteLine("Starting cell is: [" + i + ", " + j + "]");
                        return;
                    }
                }
            }


            topTableRow = -1;
        }



        /// <summary>
        /// Given the coordinates of the first data cell in the table, finds the right edge of the table by looking for its border
        /// </summary>
        /// <param name="worksheet"the worksheet we are currently working on</param>
        /// <param name="row">the row of the first data cell</param>
        /// <param name="col">the column of the first data cell</param>
        /// <returns>the column of the right edge of the table, or the specified column if the table edge isnt found</returns>
        private static int FindRightSideOfTable(ExcelWorksheet worksheet, int row, int col)
        {
            for (int j = col; j <= worksheet.Dimension.Columns; j++)
            {
                if (isEndOfTable(worksheet.Cells[row, j]))
                {
                    return j;
                }
            }

            return col; //Default: return the original cell
        }



        /// <summary>
        /// Checks if the specified cell is on the right edge of an excel table. This is determaned 
        /// by its borders.
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell is the right edge of the table, and false otherwise</returns>
        private static bool isEndOfTable(ExcelRange cell)
        {
            var border = cell.Style.Border;

            return !border.Right.Style.Equals(ExcelBorderStyle.None);
        }



        /// <summary>
        /// Given the coordinates of the first data cell in the table, finds the top row of the table by looking for its border
        /// </summary>
        /// <param name="worksheet"the worksheet we are currently working on</param>
        /// <param name="row">the row of the first data cell</param>
        /// <param name="col">the column of the first data cell</param>
        /// <returns>the top row of the table, or the specified row if the table top isnt found</returns>
        private static int FindTopEdgeOfTable(ExcelWorksheet worksheet, int row, int col)
        {
            for (int i = row; i >= 1; i--)
            {
                if (isTopOfTable(worksheet.Cells[i, col]))
                {
                    return i;
                }
            }

            return row; //Default: return the original cell
        }



        /// <summary>
        /// Checks if the specified cell is on the top row of an excel table. This is determaned 
        /// by its borders.
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell is the top row of the table, and false otherwise</returns>
        private static bool isTopOfTable(ExcelRange cell)
        {
            var border = cell.Style.Border;

            return !border.Bottom.Style.Equals(ExcelBorderStyle.None) && !border.Top.Style.Equals(ExcelBorderStyle.None);
        }




        /// <summary>
        /// Unmerges all the merged sections in the worksheet.
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private static void UnMergeMergedSections(ExcelWorksheet worksheet)
        {

            rowNeedsResize = new bool[worksheet.Dimension.End.Row];


            ExcelWorksheet.MergeCellsCollection mergedCells = worksheet.MergedCells;


            for (int i = mergedCells.Count - 1; i >= 0; i--)
            {
                var merged = mergedCells[i];

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

                case MergeType.MAIN_HEADER:
                    return false;

                case MergeType.EMPTY:
                    break;

                case MergeType.MINOR_HEADER:
                    SetMinorHeaderCellSize(worksheet, currentCells);
                    break;

                default: //If its a data cell
                    ChooseDataCellWidth(currentCells, originalStyle);
                    MarkColumnsForDeletion(worksheet, currentCells);
                    break;
            }



            currentCells.Merge = false; //unmerge range


            worksheet.Row(currentCells.Start.Row).Height = initialHeigth;



            SetCellStyles(currentCells, originalStyle); //reset the style to the way it was


            return true;

        }



        /// <summary>
        /// Counts the number of cells in a merged section
        /// </summary>
        /// <param name="mergedCells">the section to be counted</param>
        /// <returns>the number of cells in the specified section</returns>
        private static int CountNumCellsInMerge(ExcelRange mergedCells)
        {
            return mergedCells.End.Column - mergedCells.Start.Column + 1;
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

            if (isEmptyCell(cell))
            {
                return MergeType.EMPTY;
            }

            if (isDataCell(cell))
            {
                return MergeType.DATA;
            }

            if (isInsideTable(cell))
            {
                return MergeType.MINOR_HEADER;
            }


            //Otherwise is just a regular header
            return MergeType.MAIN_HEADER;
        }



        /// <summary>
        /// Checks if a cell has no text
        /// </summary>
        /// <param name="currentCells">the cell that is being checked for text</param>
        /// <returns>true if there is no text in the cell, and false otherwise</returns>
        private static bool isEmptyCell(ExcelRange currentCells)
        {
            return currentCells.Text == null || currentCells.Text.Length == 0;
        }



        /// <summary>
        /// Checks if the cell at the specified coordinates is a data cell. This is used by the
        /// current implementation: a cell is a data row if it has a $ in it.
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell is a data cell and false otherwise</returns>
        private static bool isDataCell(ExcelRange cell)
        {

            return cell.Text.StartsWith("$");

        }



        /// <summary>
        /// Checks if the specified cell is inside the table in the worksheet, and not a header 
        /// above the table
        /// </summary>
        /// <param name="cell">the cell whose location is being checked</param>
        /// <returns>true if the specified cell is inside a table and false otherwise</returns>
        private static bool isInsideTable(ExcelRange cell)
        {

            return cell.Start.Row >= topTableRow;

        }




        /// <summary>
        /// Increases the size of the specified cell (height and/or width depending on the circumstances) to have
        /// the text better fit into it
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="mergedCells">the cells that need resizing</param>
        private static void SetMinorHeaderCellSize(ExcelWorksheet worksheet, ExcelRange mergedCells)
        {

            string cellText = mergedCells.Text;

            double requiredWidth = GetCellWidthFromLargestWord(cellText, mergedCells.Style.Font.Size); //GetWidthOfCellText(cellText, mergedCells.Style.Font.Size);

            UpdateColumnDesiredWidth(mergedCells.Start.Column, requiredWidth);


            //Now resize the row if needed
            double actualWidth = worksheet.Column(mergedCells.Start.Column).Width;

            if (requiredWidth > actualWidth) //if we cant fit it all in 1 line
            {

                //mark row for resize
                rowNeedsResize[mergedCells.Start.Row - 1] = true;

            }


            //old implementation
            /* 
            
            if(cellText.Length <= SMALL_TEXT_LENGTH)
            {
                //Resize width
                UpdateColumnDesiredWidth(mergedCells.Start.Column, 
                    GetWidthOfCellText(cellText, mergedCells.Style.Font.Size));


                Console.WriteLine("cell address " + mergedCells.Address + " is getting widened to 5 characters");
                return;
            }


            double cellCapacity = GetNumCharactersThatFitInCell(mergedCells, true);
            
            //If just increasing the cell width by a small amount is sufficent
            if(cellCapacity + MINOR_WIDTH_INCREASE >= cellText.Length)
            {
                UpdateColumnDesiredWidth(mergedCells.Start.Column,
                    GetWidthOfCellText(cellText, mergedCells.Style.Font.Size));


                Console.WriteLine("cell address " + mergedCells.Address + " is getting wider by a little bit");
            }
            else
            {
                int rowNumber = mergedCells.Start.Row;

                if(worksheet.Row(rowNumber).Height <= worksheet.DefaultRowHeight) //if the row is not already taller than the default
                {
                    worksheet.Row(rowNumber).Height = worksheet.DefaultRowHeight * 2; //double row hieght
                    Console.WriteLine("cell address " + mergedCells.Address + " is getting taller");
                }
                else
                {
                    Console.WriteLine("cell address " + mergedCells.Address + " has already gotten wider");
                }
                
            }
            */

        }



        /// <summary>
        /// Figures out how many characters can be held in a cell of the specified length using 
        /// the specified font size. This method is the inverse operation of the method GetWidthOfCellText
        /// </summary>
        /// <param name="cellWidth">the width of the cell</param>
        /// <param name="fontSizeUsed">the size of the font used by text in the cell</param>
        /// <returns>the number of characters that can fit into the cell and still leave a small margin</returns>
        private static double GetNumCharactersThatFitInCell(double cellWidth, double fontSizeUsed)
        {
            double characterWidth = fontSizeUsed / DEFAULT_FONT_SIZE;
            double numCharacters = 2 + (cellWidth / characterWidth);

            return numCharacters;
        }



        /// <summary>
        /// Figures out how many characters can be held in a cell of the specified length using 
        /// the specified font size. This method is just a convienence overload of 
        /// GetNumCharactersThatFitInCell(double cellWidth, double fontSizeUsed) that lets you simply pass
        /// in the cells in question and the font size and cell width will be calculated for you.
        /// </summary>
        /// <param name="cells">the cells whose character capacity needs to be calculated</param>
        /// <param name="countOnlyFirstCell">
        /// tells the function if it should use the width of the full 
        /// cell range or just the width of the first address in the range
        /// </param>
        /// <returns>the number of characters that can fit into the cell and still leave a small margin</returns>
        private static double GetNumCharactersThatFitInCell(ExcelRange cells, bool countOnlyFirstCell)
        {
            double characterWidth = cells.Style.Font.Size / DEFAULT_FONT_SIZE;

            double cellWidth = (countOnlyFirstCell ?
                cells.Worksheet.Column(cells.Start.Column).Width :
                GetOriginalCellWidth(cells));


            double numCharacters = 2 + (cellWidth / characterWidth);

            return numCharacters;
        }




        /// <summary>
        /// Chooses the best width to use for the column containing the specified data cell and
        /// stores it in a dictionary for later, when the resize is actually done.
        /// </summary>
        /// <param name="mergedDataCells">the cells that should be used to determan the best column width</param>
        /// <param name="originalStyle">the style of the cell before it was unmerged</param>
        private static void ChooseDataCellWidth(ExcelRange mergedDataCells, ExcelStyle originalStyle)
        {
            double initialWidth = GetOriginalCellWidth(mergedDataCells);
            double properWidth = GetWidthOfCellText(mergedDataCells.Text, originalStyle.Font.Size);

            double desiredWidth = Math.Min(initialWidth, properWidth);

            //Update our dictionary with the size that this column should be.
            UpdateColumnDesiredWidth(mergedDataCells.Start.Column, desiredWidth);
        }



        /// <summary>
        /// Finds the total width of a merged cell
        /// </summary>
        /// <param name="mergedCells">the merged cells whose width must be mesured</param>
        /// <returns>the total width of the merged cells</returns>
        private static double GetOriginalCellWidth(ExcelRange mergedCells)
        {
            double width = 0;
            //iterate horizontally through every cell in the range and add its width to the total width


            for (int columnIndex = mergedCells.Start.Column; columnIndex <= mergedCells.End.Column; columnIndex++)
            {
                double columnWidth = mergedCells.Worksheet.Column(columnIndex).Width;
                width += columnWidth;
            }


            //If the cell is merged vertically we need to only count the width of 1 row.
            width /= mergedCells.Rows;


            return width;
        }



        /// <summary>
        /// Calculates a column width based on the longest word in the specified text. The remaining text is expected
        /// to wrap on other lines.
        /// </summary>
        /// <param name="columnText">the text in (one of the cells of) the column being resized</param>
        /// <param name="fontSizeUsed">the font size of the text displayed in the column</param>
        /// <returns>the column width that can hold the largest work in the cell text</returns>
        private static double GetCellWidthFromLargestWord(string columnText, double fontSizeUsed)
        {
            if (columnText == null || columnText.Length == 0)
            {
                return 0;
            }


            string[] words = columnText.Split(' ');

            int max = words[0].Length;

            for (int i = 1; i < words.Length; i++)
            {
                if (words[i].Length > max)
                {
                    max = words[i].Length;
                }
            }


            return GetWidthOfCellText(max, fontSizeUsed);
        }




        /// <summary>
        /// Calculates a column width that would be sufficent for a column that stores the specified text in a single line
        /// </summary>
        /// <param name="columnText">the text in (one of the cells of) the column being resized</param>
        /// <param name="fontSizeUsed">the font size of the text displayed in the column</param>
        /// <param name="givePadding">if true (or default) adds space for 2 extra characters in the cell with</param>
        /// <returns>the appropriate column width</returns>
        private static double GetWidthOfCellText(string columnText, double fontSizeUsed, bool givePadding = true)
        {
            int padding = (givePadding ?  2  :  0);

            double characterWidth = fontSizeUsed / DEFAULT_FONT_SIZE;

            double lengthOfText = (columnText.Length + padding) * characterWidth;

            //double lengthOfText = columnText.Length + padding; //if you want to ignore font size use this

            return lengthOfText;
        }



        /// <summary>
        /// Calculates a column width that would be sufficent for a column that stores text of the specified length in a single line.
        /// </summary>
        /// <param name="textLength">the length of the text in (one of the cells of) the column being resized</param>
        /// <param name="fontSizeUsed">the font size of the text displayed in the column</param>
        /// <returns>the appropriate column width</returns>
        private static double GetWidthOfCellText(int textLength, double fontSizeUsed)
        {

            double characterWidth = fontSizeUsed / DEFAULT_FONT_SIZE;

            double lengthOfCell = (textLength + 2) * characterWidth;

            //double lengthOfCell = textLength + 2; //if you want to ignore font size use this

            return lengthOfCell;
        }



        /// <summary>
        /// Updates the dictionary with the proper desired width of the specified column
        /// </summary>
        /// <param name="columnNumber">the column whose desired size we are updating</param>
        /// <param name="desiredSize">the desired size of the column</param>
        private static void UpdateColumnDesiredWidth(int columnNumber, double desiredSize)
        {

            if (!desiredColumnSizes.ContainsKey(columnNumber))
            {
                desiredColumnSizes.Add(columnNumber, desiredSize);
            }
            else if (desiredColumnSizes[columnNumber] < desiredSize)
            {
                desiredColumnSizes[columnNumber] = desiredSize;
            }
        }



        /// <summary>
        /// Adds the column numbers of all the empty columns that come about as a result of an unmerge - to the set of columns
        /// that are candidates for deletion
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="cells">the merge cell that will be unmerged</param>
        private static void MarkColumnsForDeletion(ExcelWorksheet worksheet, ExcelRange cells)
        {

            for (int i = cells.Start.Column + 1; i <= cells.End.Column; i++)
            {
                columnsToDelete.Add(i);
            }
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
            currentCells.Style.Font = style.Font;
            currentCells.Style.HorizontalAlignment = style.HorizontalAlignment;
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
        /// Resizes all columns in the specified worksheet to match ba the desired size as specified
        /// in the desiredColumnSizes Dictionary.
        /// </summary>
        /// <param name="worksheet">the worksheet that needs its columns resized</param>
        private static void ResizeColumns(ExcelWorksheet worksheet)
        {
            foreach (KeyValuePair<int, double> data in desiredColumnSizes)
            {
                worksheet.Column(data.Key).Width = data.Value;
            }
        }



        /// <summary>
        /// Resizes the rows that have non-data cells with insuffiecent space for thier text
        /// </summary>
        /// <param name="worksheet"></param>
        private static void ResizeRows(ExcelWorksheet worksheet) 
        {
            for(int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {

                //Check if the row still needs a resize. We might have previously made a 
                //column wider and now no longer need a resize.
                if (RowStillNeedsResize(worksheet, row))
                {
                    worksheet.Row(row).Height = worksheet.DefaultRowHeight * 2; //double row hieght
                }

            }
        }




        /// <summary>
        /// Checks if a row that has been marked for resize still needs a resize, despite the column enlargments already done.
        /// </summary>
        /// <param name="worksheet">the worksheet currently bieng cleaned</param>
        /// <param name="rowNumber">the row of the worksheet we are checking</param>
        /// <returns>true if the row has at least one cell that needs more space</returns>
        private static bool RowStillNeedsResize(ExcelWorksheet worksheet, int rowNumber)
        {

            if (!rowNeedsResize[rowNumber - 1])
            {
                return false;
            }

            if (worksheet.Row(rowNumber).Height > worksheet.DefaultRowHeight)
            {
                return false; //if its already larger than the default, we don't want to change it
            }




            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {

                ExcelRange cell = worksheet.Cells[rowNumber, col];
                string cellText = cell.Text;


                //Empty cells dont need resize, and data cells have already been resized
                if (isEmptyCell(cell) || isDataCell(cell))
                {
                    continue;
                }



                //double requiredWidth = GetCellWidthFromLargestWord(cellText, cell.Style.Font.Size);
                double requiredWidth = GetWidthOfCellText(cellText, cell.Style.Font.Size, false);
                double actualWidth = worksheet.Column(cell.Start.Column).Width;

                if (requiredWidth > actualWidth) //if we cant fit it all in 1 line
                {
                    return true;
                }
            }


            return false;
        }



        /// <summary>
        /// Deletes all columns marked for deletion if it is safe to do so
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        private static void DeleteColumns(ExcelWorksheet worksheet)
        {

            //To avoid issues with column numbers changing, the columns must be deleted in revese order
            int[] columns = columnsToDelete.ToArray<int>();
            Array.Sort(columns);


            for (int i = columns.Length - 1; i >= 0; i--)
            {
                int columnNumber = columns[i];


                if (!ColumnIsSafeToDelete(worksheet, columnNumber))
                {
                    continue;
                }



                Console.WriteLine("Column " + columnNumber + " is being deleted");

                worksheet.DeleteColumn(columnNumber);
            }
        }



        /// <summary>
        /// Checks if the specified column has any cells with text in it and is therefore unsafe
        /// to delete
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently working on</param>
        /// <param name="column">the column we want to delete</param>
        /// <returns>true if there is no text anywhere in the column and false otherwise</returns>
        private static bool ColumnIsSafeToDelete(ExcelWorksheet worksheet, int column)
        {

            for (int row = 1; row < worksheet.Dimension.Rows; row++)
            {
                string cellText = worksheet.Cells[row, column].Text;

                if (cellText != null && cellText.Length > 0)
                {
                    return false;
                }
            }

            return true;
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
