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

        //Just convienences that you can pass to this classes constructor
        public static readonly Predicate<ExcelRange> IsEmptyOrNonDataCell = 
            (cell => FormulaManager.IsEmptyCell(cell) || !FormulaManager.IsDataCell(cell));


        public static readonly Predicate<ExcelRange> IsNonDataCell =
            (cell => !FormulaManager.IsEmptyCell(cell) && !FormulaManager.IsDataCell(cell));




        private ExcelIterator iter;
        private Predicate<ExcelRange> IsBeyondFormulaRange;



        /// <summary>
        /// Constructs a FullTableFormulaGenerator that considers a forumal range to end when it encounters a 
        /// cell that is either empty or is not a data cell (does not start with a $ sign). If you want to 
        /// change this condition, pass in your own predicate to this constructor.
        /// </summary>
        public FullTableFormulaGenerator()
        {
            IsBeyondFormulaRange = IsEmptyOrNonDataCell;
        }



        /// <summary>
        /// Constructs a FullTableFormulaGenerator that considers a forumal range to end when it encounters a 
        /// cell that matches the specified predicate
        /// </summary>
        /// <param name="IsBeyondFormulaRange">a predicate that tells the system if a cell is considered outside the formula range</param>
        public FullTableFormulaGenerator(Predicate<ExcelRange> FormulaRangeEndsIf)
        {
            this.IsBeyondFormulaRange = FormulaRangeEndsIf;
        }


        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            iter = new ExcelIterator(worksheet);

            //for each header in the report that needs a formula 
            foreach (string header in headers)              
            {

                //Ensure that the header was intended for this class and not the DistantRowsFormulaGenerator class
                if (FormulaManager.IsNonContiguousFormulaRange(header))
                {
                    return;
                }


                iter.SetCurrentLocation(1, 1);
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
        private void FillInFormulas(ExcelWorksheet worksheet, int row, int col)
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
        private int FindTopRowOfFormulaRange(ExcelWorksheet worksheet, int row, int col)
        {
            ExcelIterator iterateOverFormulaRange = new ExcelIterator(iter);

            Tuple<int, int> cellAboveRange = iterateOverFormulaRange
                .GetCellCoordinates(ExcelIterator.SHIFT_UP, stopIf:IsBeyondFormulaRange)
                .Last();



            return cellAboveRange.Item1 + 1; //The row below that cell
        }

    }
}
