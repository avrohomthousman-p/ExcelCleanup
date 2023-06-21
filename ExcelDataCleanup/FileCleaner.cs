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

            UnMergeMergedSections(worksheet);

            //DeleteColumns(worksheet);

            ResizeAllColumns(worksheet);
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


            if(!currentCells.Merge)
            {
                return false;
            }


            if (IsMajorHeader(currentCells))
            {
                currentCells.Style.WrapText = false;
                currentCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                Console.WriteLine("major header at " + currentCells.Address);
            }


            //Sometimes unmerging a cell changes the row height. We need to reset it to its starting value
            double initialHeigth = worksheet.Row(currentCells.Start.Row).Height;


            //record the style we had before any changes were made
            ExcelStyle originalStyle = currentCells.Style;


            //unmerge range
            currentCells.Merge = false; 


            //restore original height
            worksheet.Row(currentCells.Start.Row).Height = initialHeigth;


            //restore the original style
            SetCellStyles(currentCells, originalStyle);


            if (IsMinorHeader(currentCells))
            {
                currentCells.Style.WrapText = false;
            }


            return true;

        }



        /// <summary>
        /// Checks if the specified cell is a major header
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the specified cell contains a major header, and false otherwise</returns>
        private static bool IsMajorHeader(ExcelRange cell)
        {
            return cell.Start.Row < firstRowOfTable;
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

            return cell.Text.StartsWith("$");

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
            for(int col = worksheet.Dimension.Columns; col >= 1;  col--)
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
        /// Checks if a column is safe to delete becuase it is empty other than possibly having major headers in it.
        /// </summary>
        /// <param name="worksheet">the worksheet where the column can be found</param>
        /// <param name="col">the column being checked</param>
        /// <returns></returns>
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
                    destinationCell.Value = originCell.Value;
                    originCell.Value = null;
                }
            }
        }




        /// <summary>
        /// Resizes all the columns in the worksheet to a size that better fits the contents
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private static void ResizeAllColumns(ExcelWorksheet worksheet)
        {
            for(int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                worksheet.Column(col).Width = ChooseBestColumnWidth(worksheet, col);
            }
        }



        /// <summary>
        /// Chooses the optimal width of the specified column.
        /// 
        /// Current implemenatation: choosewidth based on the length of the text in the first cell in that column 
        /// (that isnt a major header).
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        /// <param name="col">the column in need of resizing</param>
        /// <returns>the best width to use for the specified column</returns>
        private static double ChooseBestColumnWidth(ExcelWorksheet worksheet, int col)
        {
            ExcelRange currentCell = worksheet.Cells[firstRowOfTable, col];

            if (IsEmptyCell(currentCell))
            {
                //default: original size
                return worksheet.Column(col).Width;
            }
            else
            {
                return GetWidthOfCellText(currentCell, true);
            }

        }




        /// <summary>
        /// Calculates a cell width that would be sufficent to store the specified text in a single line
        /// </summary>
        /// <param name="cell">the cell whose text must be mesured</param>
        /// <param name="givePadding">if true (or default) adds space for 2 extra characters in the cell with</param>
        /// <returns>the appropriate column width</returns>
        private static double GetWidthOfCellText(ExcelRange cell, bool givePadding = true)
        {
            return GetWidthOfCellText(cell.Text, cell.Style.Font.Size, givePadding);
        }




        /// <summary>
        /// Calculates a cell width that would be sufficent to store the specified text in a single line
        /// </summary>
        /// <param name="columnText">the text in (one of the cells of) the column being resized</param>
        /// <param name="fontSizeUsed">the font size of the text displayed in the column</param>
        /// <param name="givePadding">if true (or default) adds space for 2 extra characters in the cell with</param>
        /// <returns>the appropriate column width</returns>
        private static double GetWidthOfCellText(string columnText, double fontSizeUsed, bool givePadding = true)
        {
            int padding = (givePadding ? 2 : 0);

            double characterWidth = fontSizeUsed / DEFAULT_FONT_SIZE;

            double lengthOfText = (columnText.Length + padding) * characterWidth;

            //double lengthOfText = columnText.Length + padding; //if you want to ignore font size use this

            return lengthOfText;
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
