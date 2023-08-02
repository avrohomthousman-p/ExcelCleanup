using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CompatableExcelCleaner
{
    /// <summary>
    /// Implementation of the IFormulaGenerator interface that searches for a row with the specifed header
    /// and adds a formula that spans as far up as it can.
    /// </summary>
    internal class FullTableFormulaGenerator : IFormulaGenerator
    {

        private ExcelIterator iter;


        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            iter = new ExcelIterator(worksheet);

            //for each header in the report that needs a formula 
            foreach (string header in headers)              
            {
                
                var allHeaderCoordinates = iter.FindAllMatchingCoordinates(cell => cell.Text == header); 
                
                //Find each instance of that header and add formulas
                foreach(var coordinates in allHeaderCoordinates)
                {
                    FillInFormulas(worksheet, coordinates.Item1, coordinates.Item2);
                }
            }

        }




        /// <summary>
        /// Inserts the formulas in each cell in the formula range that requires it.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="row">the row of the header</param>
        /// <param name="col">the column of the header</param>
        protected virtual void FillInFormulas(ExcelWorksheet worksheet, int row, int col)
        {
            iter.SetCurrentLocation(row, col);

            foreach (ExcelRange cell in iter.GetCells(ExcelIterator.SHIFT_RIGHT))
            {
                if (FormulaManager.IsEmptyCell(cell) || !FormulaManager.IsDataCell(cell))
                {
                    continue;
                }


                int topRowOfRange = FindTopRowOfFormulaRange(worksheet, row, col);

                cell.FormulaR1C1 = FormulaManager.GenerateFormula(worksheet, topRowOfRange, row - 1, iter.GetCurrentCol());
                cell.Style.Locked = true;

                Console.WriteLine("Cell " + cell.Address + " has been given this formula: " + cell.Formula);
            }

        }




        /// <summary>
        /// Given the coordinates to the bottom cell in a formula range, checks how far up the range goes
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the row number of the bottom cell in the range</param>
        /// <param name="col">the column number of the bottom cell in the range</param>
        /// <returns>the row number of the top most cell thats still part of the formula range</returns>
        protected virtual int FindTopRowOfFormulaRange(ExcelWorksheet worksheet, int row, int col)
        {
            ExcelIterator iterateOverFormulaRange = new ExcelIterator(iter);

            Tuple<int, int> cellAboveRange = iterateOverFormulaRange
                .GetCellCoordinates(ExcelIterator.SHIFT_UP, IsBeyondFormulaRange)
                .Last();


            return cellAboveRange.Item1 + 1; //The row below that cell
        }




        /// <summary>
        /// Checks if the specified cell is outside the formula range. In this implementation, that 
        /// happens when the cell is empty or is not a data cell.
        /// </summary>
        /// <param name="cell">the cell that is being checked</param>
        /// <returns>true if the specified cell is outside the formula range, and false otherwise</returns>
        protected virtual bool IsBeyondFormulaRange(ExcelRange cell)
        {
            return FormulaManager.IsEmptyCell(cell) || !FormulaManager.IsDataCell(cell);
        }
    }




    /// <summary>
    /// The same as the FullTableFormulaGenerator class except that it allows the formula range to contain empty cells
    /// </summary>
    internal class WhitespaceFriendlyFullTableGenerator : FullTableFormulaGenerator
    {


        /// <summary>
        /// Checks if the specified cell is outside the formula range. In this implementation, that 
        /// happens when the cell is not a data cell. Empty cells are part of the formula range.
        /// </summary>
        /// <param name="cell">the cell that is being checked</param>
        /// <returns>true if the specified cell is outside the formula range, and false otherwise</returns>
        protected override bool IsBeyondFormulaRange(ExcelRange cell)
        {
            return !FormulaManager.IsEmptyCell(cell) && !FormulaManager.IsDataCell(cell);
        }
    }
}
