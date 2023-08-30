using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompatableExcelCleaner.FormulaGeneration.ReportSpecificGenerators
{

    internal class VendorInvoiceReportFormulas : PeriodicFormulaGenerator
    {

        public override void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            BreakUpInvoiceTotals(worksheet);

            base.InsertFormulas(worksheet, headers);
        }



        /// <summary>
        /// Finds all cells in the report that contain the text "Invoice Total: [some amount]" and 
        /// breaks the text and number into two seperate cells.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        private void BreakUpInvoiceTotals(ExcelWorksheet worksheet)
        {
            string regex = "Invoice Total: \\$\\d{1,3}(,\\d{3})*[.]\\d\\d";

            //Invoice totals are always in the same column, so we need to find that column
            ExcelIterator iter = new ExcelIterator(worksheet);
            iter.GetFirstMatchingCell(c => FormulaManager.TextMatches(c.Text, regex));

            //move iterator up 1 to ensure the coming loop doesnt miss the first match
            iter.SetCurrentLocation(iter.GetCurrentRow() - 1, iter.GetCurrentCol());

            var cells = iter.GetCellCoordinates(ExcelIterator.SHIFT_DOWN);
            ExcelRange cell;
            foreach (Tuple<int, int> position in cells)
            {
                cell = worksheet.Cells[position.Item1, position.Item2];

                if(FormulaManager.TextMatches(cell.Text, regex))
                {
                    //TODO: split cell into seperate cells
                }
            }
        }
    }

}
