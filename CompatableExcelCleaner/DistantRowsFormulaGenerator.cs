using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CompatableExcelCleaner
{
    /// <summary>
    /// Generates formulas that add up cells from anywhere else in the worksheet. Header data should be passed
    /// to this class in this format: "headerOfFormulaCell~header1,header2,header3" where headerOfFormula cell
    /// is the header before the cell that needs the formula and the other comma seperated headers are headers in front
    /// of cells that should be included in the sum.
    /// 
    /// Note: this class will NOT do all the formulas necessary on the worksheet, only the ones that cant be done by
    /// other systems becuase their cells are not near each other. This class should be used in addition to whatever other
    /// formula generator is appropriate for the report being cleaned.
    /// </summary>
    internal class DistantRowsFormulaGenerator
    {

        /// <summary>
        /// Adds all formulas to the worksheet as specified by the metadata in the headers array
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="headers">headers to look for to tell us which cells to add up</param>
        public static void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            foreach(string header in headers)
            {
                //Ensure that the header was intended for this class and not the FormulaGenerator
                if (!FormulaManager.IsNonContiguousFormulaRange(header))
                {
                    continue;
                }
                


                int indexOfEqualsSign = header.IndexOf('~');
                string formulaHeader = header.Substring(0, indexOfEqualsSign);
                string[] dataCells = header.Substring(indexOfEqualsSign + 1).Split(',');


                FillInFormulas(worksheet, formulaHeader, dataCells);
            }
        }



        /// <summary>
        /// Inserts a formula into the cell with the specified header.
        /// </summary>
        /// <param name="worksheet">the worksheet being given formulas</param>
        /// <param name="formulaHeader">the text that should be found near the cell requiring a formula</param>
        /// <param name="dataCells">headers pointing to cells that should be included in the formula</param>
        private static void FillInFormulas(ExcelWorksheet worksheet, string formulaHeader, string[] dataCells)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);

            ExcelRange formulaCell = iter.GetFirstMatchingCell(cell => cell.Text == formulaHeader);

            if(formulaCell == null)
            {
                Console.WriteLine("Cell with text " + formulaHeader + " not found. Formula insertion failed.");
                return;
            }


            //Now get all the addresses of the data cells that should be part of the formula
            iter.SkipWhile(ExcelIterator.SHIFT_RIGHT, cell => FormulaManager.IsEmptyCell(cell) || !FormulaManager.IsDataCell(cell));
            formulaCell = iter.GetCurrentCell();

            int dataColumn = iter.GetCurrentCol();

            int[] dataRows = GetRowsToIncludeInFormula(worksheet, dataCells);



            //now build the formula
            StringBuilder formula = new StringBuilder("SUM(");

            foreach(int i in dataRows)
            {
                formula.Append(GetAddress(worksheet, i, dataColumn)).Append(",");
            }

            formula.Remove(formula.Length - 1, 1); //delete the trailing comma

            formula.Append(")");



            //now add the formula to the cell
            formulaCell.FormulaR1C1 = formula.ToString();
            formulaCell.Style.Locked = true;

            Console.WriteLine("Cell " + formulaCell.Address + " has been given this formula: " + formulaCell.Formula);
        }



        /// <summary>
        /// Gets all the row numbers of the cells that are to be included in the formula
        /// </summary>
        /// <param name="worksheet">the worksheet that is being given formulas</param>
        /// <param name="headers">the text that signals that this data cell should be part of the formula</param>
        /// <returns>an array of row numbers of the cells that should be part of the formula</returns>
        private static int[] GetRowsToIncludeInFormula(ExcelWorksheet worksheet, string[] headers)
        {
            HashSet<string> allHeaders = new HashSet<string>(headers);

            ExcelIterator iter = new ExcelIterator(worksheet);
            return iter.FindAllMatchingCoordinates(cell => allHeaders.Contains(cell.Text))
                                .Select(tup => tup.Item1).ToArray();

        }




        /// <summary>
        /// Gets the address of a cell as it would be displayed in a formula
        /// </summary>
        /// <param name="worksheet">the worksheet the cell is in</param>
        /// <param name="row">the row the cell is in</param>
        /// <param name="col">the column the cell is in</param>
        /// <returns>the cell address</returns>
        private static string GetAddress(ExcelWorksheet worksheet, int row, int col)
        {
            return worksheet.Cells[row, col].Address;
        }
    }
}
