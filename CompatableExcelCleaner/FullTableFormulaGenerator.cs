using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        private IEnumerable<Tuple<int, int>> FindAllHeaders(ExcelWorksheet worksheet, string targetHeader)
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
        private void FillInFormulas(ExcelWorksheet worksheet, int row, int col)
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

            ExcelRange cell;

            for(row--; row >= 1; row--)
            {
                cell = worksheet.Cells[row, col];

                if(FormulaManager.IsEmptyCell(cell) || !FormulaManager.IsDataCell(cell))
                {
                    return row + 1; //Return the row just before that non-data cell
                }
            }

            return 1;
        }
    }
}
