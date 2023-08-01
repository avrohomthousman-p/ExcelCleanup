using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace CompatableExcelCleaner
{
    /// <summary>
    /// Implementation of the IFormulaGenerator interface that searches for a row with the specifed header
    /// and adds a formula that spans as far up as it can.
    /// </summary>
    internal class FullTableFormulaGenerator : IFormulaGenerator
    {
        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {


            //for each header in the report that needs a formula 
            foreach (string header in headers)              
            {

                var allHeaderCoordinates = FindAllHeaders(worksheet, header);

                //Find each instance of that header and add formulas
                foreach(var coordinates in allHeaderCoordinates)
                {
                    FillInFormulas(worksheet, coordinates.Item1, coordinates.Item2);
                }
            }

        }



        /// <summary>
        /// Finds all cells in the table that have the specified header and therefore signal the need for formulas
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="targetHeader">the text that signals the need for a formula on that row</param>
        /// <returns>the row and column of the cell with the target header as a tuple</returns>
        protected virtual IEnumerable<Tuple<int, int>> FindAllHeaders(ExcelWorksheet worksheet, string targetHeader)
        {
            ExcelRange cell;


            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    cell = worksheet.Cells[row, col];

                    if (cell.Text == targetHeader)
                    {
                        yield return new Tuple<int, int>(row, col);

                        col = 1;
                        row++;
                    }
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

            ExcelRange cell;


            //Often there are multiple columns that require a formula, so we need to iterate
            //and apply the formulas in many columns
            for (col++; col <= worksheet.Dimension.End.Column; col++)
            {

                cell = worksheet.Cells[row, col];


                if (FormulaManager.IsEmptyCell(cell))
                {
                    continue;
                }


                int topRowOfRange = FindTopRowOfFormulaRange(worksheet, row, col);

                cell.FormulaR1C1 = FormulaManager.GenerateFormula(worksheet, topRowOfRange, row-1, col);
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

            ExcelRange cell;

            for(row--; row >= 1; row--)
            {
                cell = worksheet.Cells[row, col];

                if(IsBeyondFormulaRange(cell))
                {
                    return row + 1; //Return the row just before that non-data cell
                }
            }

            return 1;
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
