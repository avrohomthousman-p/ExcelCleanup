﻿using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompatableExcelCleaner
{

    public delegate bool IsSummaryCell(ExcelRange cell);


    /// <summary>
    /// Adds formulas to the end of "sections" found inside data columns of the worksheet. A section is defined as
    /// a series of data cells that all corrispond to a single "key" which appears on the top left of that section.
    /// The first string in the list of arguments for this class should follow this pattern: r=[insert regex] 
    /// Where the regex is used to find the key for each section. After that, the titles of each data column 
    /// that need formulas should be passed in as well.
    /// </summary>
    internal class PeriodicFormulaGenerator : IFormulaGenerator
    {

        private IsDataCell isDataCell = new IsDataCell(FormulaManager.IsDollarValue); //default implementation
        private IsSummaryCell isSummaryCell = new IsSummaryCell(
            (cell => FormulaManager.IsDollarValue(cell) && cell.Style.Font.Bold)); //defualt implementation




        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.isDataCell = isDataCell;
        }



        public void SetSummaryCellDefenition(IsSummaryCell summaryCellDef)
        {
            this.isSummaryCell = summaryCellDef;
        }



        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            if (!headers[0].StartsWith("r="))
            {
                throw new ArgumentException("The argument to this formula generator must specify a regex that matches the key cell of each section");
            }



            string keyRegex = headers[0].Substring(2);

            for (int i = 1; i < headers.Length; i++)
            {
                //Ensure that the header was intended for this class and not the DistantRowsFormulaGenerator class
                if (FormulaManager.IsNonContiguousFormulaRange(headers[i]))
                {
                    continue;
                }

                InsertFormulaForHeader(worksheet, keyRegex, headers[i]);
            }
        }




        /// <summary>
        /// Adds all formulas to cells marked with the specified header
        /// </summary>
        /// <param name="worksheet">the worksheet being given formulas</param>
        /// <param name="key">a regex defining what the "key" for a data section should look like</param>
        /// <param name="targetHeader">the text the header should have</param>
        private void InsertFormulaForHeader(ExcelWorksheet worksheet, string key, string targetHeader)
        {
            var coordinates = FindStartOfDataColumn(worksheet, targetHeader);
            int row = coordinates.Item1 + 1; //start by the row after the column header
            int dataCol = coordinates.Item2;

            for(; row <= worksheet.Dimension.End.Row; row++)
            {

                FindNextKey(worksheet, key, ref row);

                ProcessFormulaRange(worksheet, ref row, dataCol);
                
            }
        }



        /// <summary>
        /// Finds the coordinates of the cell with the specified columnHeader
        /// </summary>
        /// <param name="worksheet">the worksheet getting formulas</param>
        /// <param name="columnHeader">the text that signaling that this is the cell we want</param>
        /// <returns>a tuple with the row and column of the cell containing the specifed text, or null if its not found</returns>
        private Tuple<int, int> FindStartOfDataColumn(ExcelWorksheet worksheet, string columnHeader)
        {
            ExcelRange cell;

            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    cell = worksheet.Cells[row, col];

                    if (FormulaManager.TextMatches(cell.Text, columnHeader))
                    {
                        return new Tuple<int, int>(row, col);
                    }
                }
            }

            return null;
        }



        /// <summary>
        /// Finds the next cell, below the specified row, that contains a key (meaning it might require a formula of its own).
        /// After this function completes, the row variable will either be pointing to the next cell with a key, or the last row
        /// plus 1, if there is no next key.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="key">the pattern that a key must match</param>
        /// <param name="row">the row to start searching on</param>
        private void FindNextKey(ExcelWorksheet worksheet, string key, ref int row)
        {
            ExcelRange cell;

            for(; row <= worksheet.Dimension.End.Row; row++)
            {
                cell = worksheet.Cells[row, 1];

                if(FormulaManager.TextMatches(cell.Text, key))
                {
                    return;
                }
            }
        }




        /// <summary>
        /// Finds the bounds of the formula range and does the actual insertion of the formula
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the row number of the key for the section we are processing</param>
        /// <param name="dataCol">the column we should look for summary cells in</param>
        private void ProcessFormulaRange(ExcelWorksheet worksheet, ref int row, int dataCol)
        {

            int start = row; //the formula range starts here, at the first non-empty cell



            row++; //the first cell has the key so it isnt empty, and causes the skip to end immideatly
            SkipEmptyCells(worksheet, ref row, 1);



            //SkipEmptyCells leaves the row variable referencing the first non empty cell found, which
            // is at the start of the next section. We want it at the last cell of this section.
            row--; 



            //Ensure there is a summary cell (some sections dont have one)
            int summaryRow = FindSummaryCellRow(worksheet, row, start, dataCol);
            if (summaryRow == -1)
            {
                return; //no summary cell
            }



            //Insert formulas
            ExcelRange summaryCell = worksheet.Cells[summaryRow, dataCol];
            Console.WriteLine("adding formula to " + summaryCell.Address);
            summaryCell.Formula = FormulaManager.GenerateFormula(worksheet, start, summaryRow - 1, dataCol);
            summaryCell.Style.Locked = true;
        }



        /// <summary>
        /// Advances the row variable to reference the first non empty cell from itself or below
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the first row in the column to check</param>
        /// <param name="col">the column we are scanning</param>
        private void SkipEmptyCells(ExcelWorksheet worksheet, ref int row, int col)
        {
            ExcelRange cell = worksheet.Cells[row, col];

            while (FormulaManager.IsEmptyCell(cell) && row + 1 <= worksheet.Dimension.End.Row)
            {
                row++;
                cell = worksheet.Cells[row, col];
            }
        }



        /// <summary>
        /// Finds the lowest summary cell between the specified bottom and top row and in the specified column
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="bottomRow">the row we should start chacking at</param>
        /// <param name="topRow">the upper row limit that the formula range cannot go past</param>
        /// <param name="col">the column to look in</param>
        /// <returns>the row number of the summary cell found, or -1 if there is no summary cell</returns>
        private int FindSummaryCellRow(ExcelWorksheet worksheet, int bottomRow, int topRow, int col)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, bottomRow, col);
            foreach (ExcelRange cell in iter.GetCells(ExcelIterator.SHIFT_UP, cell => cell.Start.Row < topRow))
            {
                if (isSummaryCell(cell))
                {
                    return cell.Start.Row;
                }
            }


            return -1; //no summary cell found
        }




        /// <summary>
        /// Checks if the specified cell is the first cell to come after the formula range. In this implemenatation, a cell is 
        /// the last in the range if it contains a top border. This is just a utility method that can be used passed to the method
        /// SetSummaryCellDefenition.
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell is the last cell in the formula range, and false otherwise</returns>
        public bool HasTopBorder(ExcelRange cell)
        {
            var border = cell.Style.Border;

            return !border.Top.Style.Equals(ExcelBorderStyle.None) && border.Bottom.Style.Equals(ExcelBorderStyle.None);
        }
    }
}
