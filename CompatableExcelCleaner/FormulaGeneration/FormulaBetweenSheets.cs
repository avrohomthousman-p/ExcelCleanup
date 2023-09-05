using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompatableExcelCleaner.FormulaGeneration
{
    /// <summary>
    /// Implementation of IFormulaGenerator that builds formulas that sum up the totals at the bottom of multiple worksheets.
    /// The header argments for this class should simply be all the sheet numbers that are to be included in the sum, expressed 
    /// in the format "sheet[sheetNum]" with no brackets. The sheet numbers are zero based.
    /// </summary>
    internal class FormulaBetweenSheets : IFormulaGenerator
    {
        private IsDataCell dataCellDef = new IsDataCell(FormulaManager.IsDollarValue);



        public void InsertFormulas(ExcelWorksheet mainWorksheet, string[] headers)
        {
            bool[] isNegative = headers.Select(s => s.StartsWith("-")).ToArray();
            int[] sheetsToAdd = headers.Select(s => Int32.Parse(s.Substring(5))).ToArray();

            ExcelWorkbook workbook = mainWorksheet.Workbook;

            int summaryCellNum = 0; //tracks which summary cell we are currently working on
            ExcelIterator mainIterator = new ExcelIterator(mainWorksheet, mainWorksheet.Dimension.End.Row, 1);
            foreach(ExcelRange cell in mainIterator.GetCells(ExcelIterator.SHIFT_RIGHT))
            {
                if (!dataCellDef(cell))
                {
                    continue;
                }

                summaryCellNum++;
                cell.Formula = "TODO";
                cell.Style.Locked = true;
            }
        }



        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.dataCellDef = isDataCell;
        }
    }
}
