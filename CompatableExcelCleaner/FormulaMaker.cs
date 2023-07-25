﻿using OfficeOpenXml;
using System;
using System.IO;
using System.Collections.Generic;


namespace CompatableExcelCleaner
{
    /// <summary>
    /// Replaces static values in excel files with formulas that will change when the data is updated
    /// </summary>
    public class FormulaMaker
    {


        /// <summary>
        /// Adds all necissary formulas to the appropriate cells in the specified file
        /// </summary>
        /// <param name="sourceFile">the excel file needing formulas, stored as an array/stream of bytes</param>
        /// <param name="reportName">the name of the report</param>
        /// <returns>the byte stream/arrray of the modified file</returns>
        public static byte[] AddFormulas(byte[] sourceFile, string reportName)
        {

            using (ExcelPackage package = new ExcelPackage(new MemoryStream(sourceFile)))
            {
                ExcelWorksheet worksheet;
                for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
                {
                    worksheet = package.Workbook.Worksheets[i];

                    //TODO: add formulas
                    var ranges = GetRowRangeForFormula(worksheet, "Income");
                    foreach (var item in ranges)
                    {
                        Console.WriteLine("range from " + item.Item1 + " to " + item.Item2);
                    }

                }
            }

            return sourceFile;
        }



        /// <summary>
        /// Gets the row numbers of the first and last rows that should be included in the formula
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="targetText">the text to look for to signal the start and end row</param>
        /// <returns>a tuple containing the start-row and end-row of the formula range</returns>
        private static IEnumerable<Tuple<int, int>> GetRowRangeForFormula(ExcelWorksheet worksheet, string targetText)
        {
            ExcelRange cell;


            for(int row = 1; row < worksheet.Dimension.End.Row; row++)
            {
                for(int col = 1; col < worksheet.Dimension.End.Column; col++)
                {
                    cell = worksheet.Cells[row, col];

                    if (cell.Text == targetText)
                    {
                        //search for end of sequence
                        int end = FindEndOfFormulaRange(worksheet, row, col, "Total " + targetText);

                        if (end > 0)
                        {
                            yield return new Tuple<int, int>(row, end);

                            //for the next iteration, jump to after the formula range we just returned
                            row = end + 1;
                            col = 1;
                        }
                    }
                }
            }
        }




        /// <summary>
        /// Given the Cell coordinates of the starting cell in a formula range, finds the ending cell for 
        /// that range.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="row">the row number of the starting cell in the formula range</param>
        /// <param name="col">the column number of the starting cell in the formula range</param>
        /// <param name="targetText">the text to look for that signals the end cell of the formula range</param>
        /// <returns>the row number of the last cell in the formula range, or -1 if no appropriate last cell is found</returns>
        private static int FindEndOfFormulaRange(ExcelWorksheet worksheet, int row, int col, string targetText)
        {
            ExcelRange cell;

            for(int i = row + 1; i < worksheet.Dimension.End.Row; i++)
            {
                cell = worksheet.Cells[i, col];
                if (cell.Text == targetText)
                {
                    return i;
                }
            }


            return -1;
        }




        /// <summary>
        /// Inserts the formulas in each cell in the formula range that requires it.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="startRow">the first row of the formula range (containing the header)</param>
        /// <param name="endRow">the last row of the formula range (containing the total)</param>
        /// <param name="col">the column of oth the header and total for the formula range</param>
        private static void FillInFormulas(ExcelWorksheet worksheet, int startRow, int endRow, int col)
        {

            //First, skip all empty cells between the current column and the actual data columns
            ExcelRange cell;

            for(col++; col <= worksheet.Dimension.Columns; col++)
            {
                cell = worksheet.Cells[endRow, col];
                if (IsNumeric(cell.Text))
                {
                    break;
                }
            }


            //If the loop ended becuase we went out of bounds
            if(col > worksheet.Dimension.End.Column)
            {
                return;
            }


            for(; col <= worksheet.Dimension.Columns; col++)
            {
                cell = worksheet.Cells[endRow, col];
                if (IsNumeric(cell.Text))
                {
                    cell.FormulaR1C1 = GenerateFormula(worksheet, startRow, endRow, col);
                }
                else
                {
                    return;
                }
            }

        }




        /// <summary>
        /// Checks if a string can be converted into a double
        /// </summary>
        /// <param name="data">the text being checked</param>
        /// <returns>true if the text can be safely converted to a double and false otherwise</returns>
        private static bool IsNumeric(string data)
        {
            double unused;

            return Double.TryParse(data, out unused);
        }



        /// <summary>
        /// Generates the formula for the cells in the given range
        /// </summary>
        /// <param name="worksheet">the worksheet currently getting formulas</param>
        /// <param name="startRow">the starting row of the formula range</param>
        /// <param name="endRow">the ending row of the formula range</param>
        /// <param name="col">the column the formula is for</param>
        /// <returns>the proper formula for the specified formula range</returns>
        private static string GenerateFormula(ExcelWorksheet worksheet, int startRow, int endRow, int col)
        {
            ExcelRange cells = worksheet.Cells[startRow - 1, col, endRow + 1, col];

            return "SUM(" + cells.Address + ")";
        }

    }
}