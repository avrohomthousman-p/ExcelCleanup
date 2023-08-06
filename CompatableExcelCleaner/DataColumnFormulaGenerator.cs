using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompatableExcelCleaner
{

    /// <summary>
    /// Adds formulas to the end of "sections" found inside data columns of the worksheet. A section is defined as
    /// a series of data cells seperated from the rest of the data cells by at least 
    /// one empty cell on the top and a bottom border on bottom. The titles of each data column that need formulas 
    /// should be passed in via the headers string array.
    /// </summary>
    internal class DataColumnFormulaGenerator : IFormulaGenerator
    {

        private IsDataCell isDataCell = new IsDataCell(FormulaManager.IsDollarValue); //default implementation


        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.isDataCell = isDataCell;
        }



        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            foreach (string header in headers)
            {
                InsertFormulaForHeader(worksheet, header);
            }
        }




        /// <summary>
        /// Adds all formulas to cells marked with the specified header
        /// </summary>
        /// <param name="worksheet">the worksheet being given formulas</param>
        /// <param name="targetHeader">the text the header should have</param>
        private void InsertFormulaForHeader(ExcelWorksheet worksheet, string targetHeader)
        {
            var coordinates = FindStartOfDataColumn(worksheet, targetHeader);
            int row = coordinates.Item1 + 1; //start by the row after the column header
            int col = coordinates.Item2;

            
            for(; row <= worksheet.Dimension.End.Row; row++)
            {

                SkipEmptyCells(worksheet, ref row, col);


                //Then keep moving untill you reach the end of the formula range
                //if you hit an empty cell, continue outer loop
                ProcessFormulaRange(worksheet, ref row, col);
                
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

                    if (cell.Text == columnHeader)
                    {
                        return new Tuple<int, int>(row, col);
                    }
                }
            }

            return null;
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

            while (FormulaManager.IsEmptyCell(cell) && row+1 <= worksheet.Dimension.End.Row)
            {
                row++;
                cell = worksheet.Cells[row, col];
            } 
        }




        /// <summary>
        /// Finds the bounds of the formula range and does the actual insertion of the formula
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="row">the row we should start from</param>
        /// <param name="col">the column we should look in</param>
        private void ProcessFormulaRange(ExcelWorksheet worksheet, ref int row, int col)
        {

            int start = row; //the formula range starts here, at the first non-empty cell

            ExcelRange cell;

            while (row <= worksheet.Dimension.End.Row)
            {
                cell = worksheet.Cells[row, col];

                if (IsFirstCellAfterFormulaRange(cell))
                {
                    /*
                     * NOTE: the code here is built specifically to work for the first worksheet in the report
                     * "OutstandingBalanceReport". In the event that this system works for another report as well,
                     * this code may require modification.
                     */


                    //Add a formula
                    ExcelRange oldTotalCell = worksheet.Cells[row, col + 1];
                    cell = worksheet.Cells[row, col];
                    oldTotalCell.CopyStyles(cell);
                    cell.FormulaR1C1 = FormulaManager.GenerateFormula(worksheet, start, row - 1, col);
                    cell.Style.Locked = true;

                    oldTotalCell.SetCellValue(0, 0, ""); 
                    return;
                }
                else if (FormulaManager.IsEmptyCell(cell) || !this.isDataCell(cell))
                {
                    //This isnt an actual formula range
                    return;
                }
                else
                {
                    //this is a data cell in the midle of the formula range, so keep moving
                    row++;
                }
            }

        }




        /// <summary>
        /// Checks if the specified cell is the first cell to come after the formula range. In this implemenatation, a cell is 
        /// the last in the range if it contains a top border.
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell is the last cell in the formula range, and false otherwise</returns>
        private bool IsFirstCellAfterFormulaRange(ExcelRange cell)
        {
            var border = cell.Style.Border;

            return !border.Top.Style.Equals(ExcelBorderStyle.None) && border.Bottom.Style.Equals(ExcelBorderStyle.None);
        }
    }
}
