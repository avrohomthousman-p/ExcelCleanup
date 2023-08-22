using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompatableExcelCleaner.FormulaGeneration.ReportSpecificGenerators
{

    internal delegate bool IsOutsideFormula(ExcelRange cell);


    /// <summary>
    /// Implementation of IFormulaGenerator that gives formulas to a column that is the sum of all
    /// columns to its left. This is similar to the SummaryColumn Formula generator, except that it adds
    /// all columns to the left, instead of just adding specific columns. Also, columns cannot be made negetive.
    /// </summary>
    internal class FullTableSummaryColumn : IFormulaGenerator
    {
        private IsDataCell dataCellDef = new IsDataCell(cell => FormulaManager.IsDollarValue(cell));
        private IsOutsideFormula outsideFormula = new IsOutsideFormula(cell => !FormulaManager.IsDollarValue(cell));



        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            foreach(string header in headers)
            {
                Tuple<int, int> headerCellCoords = FindHeaderCell(worksheet, header);
                AddFormulas(worksheet, headerCellCoords.Item1, headerCellCoords.Item2);
            }
        }




        /// <summary>
        /// Finds the cell that matches the specified header (and is therefore the column that needs formulas)
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="header">the text that the header cell must match</param>
        /// <returns>the row and column (as a tuple) of the header cell at the top of the formula column</returns>
        private Tuple<int, int> FindHeaderCell(ExcelWorksheet worksheet, string header)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);
            return iter.FindAllMatchingCoordinates(cell => FormulaManager.TextMatches(cell.Text, header)).First();
        }



        /// <summary>
        /// Gives each cell in the specifeid column a formula if needed
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the row number of the header of the column getting formulas</param>
        /// <param name="col">the column getting formulas</param>
        private void AddFormulas(ExcelWorksheet worksheet, int row, int col)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, row + 1, col);

            var summaryCells = iter.GetCells(ExcelIterator.SHIFT_DOWN, cell => !dataCellDef(cell));//SAFE TO MAKE CHANGE HERE

            foreach (ExcelRange cell in summaryCells)
            {
                int startColumn = GetFormulaStartColumn(worksheet, cell.Start.Row, col);
                cell.Formula = BuildFormula(worksheet, cell.Start.Row, startColumn, col - 1);
                cell.Style.Locked = true;
            }
        }




        /// <summary>
        /// Finds the column number of the first column in this formula range.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the row this formula is for</param>
        /// <param name="startCol">the column we should start iterating from</param>
        /// <returns>the column number of the leftmost column in the formula</returns>
        private int GetFormulaStartColumn(ExcelWorksheet worksheet, int row, int startCol)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, row, startCol);
            var lastCell = iter.GetCells(ExcelIterator.SHIFT_LEFT, cell => outsideFormula(cell)).Last();

            //SAFE TO MAKE CHANGE
            //The iterator returns the cell that made the predicate true, and is outside the formula range.
            //Therefore  we need to ensure that if thats what happened, we return the cell before it
            if (outsideFormula(lastCell))
            {
                return lastCell.End.Column + 1;
            }
            else
            {
                return lastCell.End.Column;
            }
        }




        /// <summary>
        /// Builds a formula that spans the horizontal area between the specified start and end columns (inclusive)
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the row this formula is for</param>
        /// <param name="startCol">the start (leftmost) column of the formula</param>
        /// <param name="endCol">the end (rightmost) column of the formula</param>
        /// <returns>a string with the proper formula to sum up the specified range</returns>
        private string BuildFormula(ExcelWorksheet worksheet, int row, int startCol, int endCol)
        {
            if(startCol > endCol)
            {
                return null; //dont insert a formula
            }

            ExcelRange formulaRange = worksheet.Cells[row, startCol, row, endCol];

            return "SUM(" + formulaRange.Address + ")";
        }




        public void SetOutsideFormulaDefenition(IsOutsideFormula isOutsideFormula)
        {
            this.outsideFormula = isOutsideFormula;
        }



        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.dataCellDef = isDataCell;
        }
    }
}
