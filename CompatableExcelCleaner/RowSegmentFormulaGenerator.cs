﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompatableExcelCleaner
{
    /// <summary>
    /// Implementation of IFormulaGenerator that looks for the specifed header followed by the other specified 
    /// (e.g. "Income" and "Total Income"), and treats the rows between those headers as a "formula range" that 
    /// gets its own formula.
    /// </summary>
    internal class RowSegmentFormulaGenerator : IFormulaGenerator
    {
        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            string startHeader, endHeader;

            foreach (string header in headers)              //for each header in the report that needs a formula 
            {

                int seperator = header.IndexOf('=');

                startHeader = header.Substring(0, seperator);
                endHeader = header.Substring(seperator + 1);

                var ranges = GetRowRangeForFormula(worksheet, startHeader, endHeader);

                foreach (var item in ranges)                // for each instance of that header
                {
                    FillInFormulas(worksheet, item.Item1, item.Item2, item.Item3);
                }
            }

        }




        /// <summary>
        /// Gets the row numbers of the first and last rows that should be included in the formula
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="startHeader">the text to look for to signal the start row</param>
        /// <param name="endHeader">The text to look for to signal the end row</param>
        /// <returns>a tuple containing the start-row, end-row, and column of the formula range</returns>
        private static IEnumerable<Tuple<int, int, int>> GetRowRangeForFormula(ExcelWorksheet worksheet, string startHeader, string endHeader)
        {
            ExcelRange cell;


            for (int row = 1; row < worksheet.Dimension.End.Row; row++)
            {
                for (int col = 1; col < worksheet.Dimension.End.Column; col++)
                {
                    cell = worksheet.Cells[row, col];

                    if (cell.Text == startHeader)
                    {

                        //search for end of sequence
                        int end = FindEndOfFormulaRange(worksheet, row, col, endHeader);

                        if (end > 0)
                        {
;
                            yield return new Tuple<int, int, int>(row, end, col);

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

            for (int i = row + 1; i <= worksheet.Dimension.End.Row; i++)
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
        /// <param name="col">the column of the header and total for the formula range</param>
        private static void FillInFormulas(ExcelWorksheet worksheet, int startRow, int endRow, int col)
        {

            ExcelRange cell;



            //Often there are multiple columns that require a formula, so we need to iterate
            //and apply the formulas in many columns
            for (col++; col <= worksheet.Dimension.End.Column; col++)
            {
                cell = worksheet.Cells[endRow, col];

                if (FormulaManager.IsDataCell(cell))
                {
                    startRow += CountEmptyCellsOnTop(worksheet, startRow, endRow, col); //Skip the whitespace on top
                    cell.FormulaR1C1 = FormulaManager.GenerateFormula(worksheet, startRow, endRow - 1, col);
                    cell.Style.Locked = true;
                    Console.WriteLine("Cell " + cell.Address + " has been given this formula: " + cell.Formula);
                }
                else if (!FormulaManager.IsEmptyCell(cell))
                {
                    return;
                }
            }

        }



        /// <summary>
        /// Counts the number of empty cells between the start header(inclusive) and the actual data cells in the 
        /// formula range.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="startRow">the row where the start header was found</param>
        /// <param name="endRow">te row where the end header was found</param>
        /// <param name="col">the column of the formula range</param>
        /// <returns>the number of empty cells at the start of the formula range</returns>
        private static int CountEmptyCellsOnTop(ExcelWorksheet worksheet, int startRow, int endRow, int col)
        {
            int emptyCells = 0;
            ExcelRange cell;

            for(;startRow <= endRow; startRow++)
            {
                cell = worksheet.Cells[startRow, col];
                if (FormulaManager.IsEmptyCell(cell))
                {
                    emptyCells++;
                }
                else
                {
                    break;
                }
            }


            return emptyCells;
        }
    }
}
