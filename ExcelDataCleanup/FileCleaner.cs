﻿using OfficeOpenXml;
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

        private static readonly int DEFAULT_COLUMN_WIDTH = 8;


        private static int topTableRow;




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

            DeleteColumns(worksheet);

            //ResizeColumns(worksheet);
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
                    if (IsDataCell(worksheet.Cells[i, j]))
                    {
                        j = FindRightSideOfTable(worksheet, i, j);
                        i = FindTopEdgeOfTable(worksheet, i, j);
                        topTableRow = i;
                        Console.WriteLine("Starting cell is: [" + i + ", " + j + "]");
                        return;
                    }
                }
            }


            topTableRow = 1; //Default is the first row
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
                if (IsEndOfTable(worksheet.Cells[row, j]))
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
        private static bool IsEndOfTable(ExcelRange cell)
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
                if (IsTopOfTable(worksheet.Cells[i, col]))
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
        private static bool IsTopOfTable(ExcelRange cell)
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


            return true;

        }



        /// <summary>
        /// Checks if the specified cell is a major header
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the specified cell contains a major header, and false otherwise</returns>
        private static bool IsMajorHeader(ExcelRange cell)
        {
            return cell.Start.Row < topTableRow;
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

            return cell.Start.Row >= topTableRow;

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
            for (int row = topTableRow; row <= worksheet.Dimension.Rows; row++)
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
            for(int row = 1; row < topTableRow; row++)
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
