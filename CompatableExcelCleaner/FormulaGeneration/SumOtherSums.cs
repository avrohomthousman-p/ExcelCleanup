﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompatableExcelCleaner.FormulaGeneration
{
    /// <summary>
    /// Functions the same way as the full table formula generator except that it adds all non-formula cells in the
    /// column instead of all cells.
    /// </summary>
    internal class SumOtherSums : FullTableFormulaGenerator
    {

        public SumOtherSums()
        {
            //in this version we override the defualt behavior set in the superclass
            this.beyondFormulaRange = new IsBeyondFormulaRange(cell => !FormulaManager.IsEmptyCell(cell) && !isDataCell(cell));
        }




        protected override void FillInFormulas(ExcelWorksheet worksheet, int row, int col)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, row, col + 1);

            foreach (ExcelRange cell in iter.GetCells(ExcelIterator.SHIFT_RIGHT))
            {
                if (HasNoTextOrFormulas(cell))
                {
                    continue;
                }


                cell.Formula = BuildFormula(worksheet, iter.GetCurrentRow(), iter.GetCurrentCol());
                cell.Formula = "Kwende";
                cell.Style.Locked = true;

                Console.WriteLine("Cell " + cell.Address + " has been given this formula: " + cell.Formula);
            }
        }



        /// <summary>
        /// Checks if the specified cell has niether text nor formulas in it and is therefore empty.
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell has no formulas and no text, or false if it has either one.</returns>
        protected bool HasNoTextOrFormulas(ExcelRange cell)
        {
            return FormulaManager.IsEmptyCell(cell) && (cell.Formula == null || cell.Formula.Length == 0);
        }



        /// <summary>
        /// Builds a formula that adds all other formulas in the formula range
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="headerRow">the row number of the formula header</param>
        /// <param name="headerCol">the column number of the formula header</param>
        /// <returns>a formula to be inserted in the proper cell</returns>
        protected virtual string BuildFormula(ExcelWorksheet worksheet, int headerRow, int headerCol)
        {
            StringBuilder result = new StringBuilder("SUM(");

            int top;
            int bottom = headerRow - 1;
            ExcelIterator iter = new ExcelIterator(worksheet, headerRow - 1, headerCol);
            foreach(ExcelRange cell in iter.GetCells(ExcelIterator.SHIFT_UP, cell => this.beyondFormulaRange(cell)))
            {

                if (FormulaManager.CellHasFormula(cell))
                {
                    top = cell.Start.Row + 1;

                    //remember that smaller row numbers are higher up.
                    //this if statement ensures that we skip all sections that are only 1 row in height
                    if(top < bottom)
                    {
                        ExcelRange range = worksheet.Cells[top, headerCol, bottom, headerCol];
                        result.Append(range.Address);
                        result.Append(",");
                    }
                    

                    bottom = cell.Start.Row - 1;
                }
            }


            //After the loop terminates, the last segment still need to be added
            top = iter.GetCurrentRow();

            if (top < bottom)
            {
                ExcelRange range = worksheet.Cells[top, headerCol, bottom, headerCol];
                result.Append(range.Address);
            }
            else
            {
                result.Remove(result.Length - 1, 1);
            }




            result.Append(")");
            return result.ToString();
        }
    }
}
