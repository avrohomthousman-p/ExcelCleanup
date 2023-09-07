using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompatableExcelCleaner.FormulaGeneration.ReportSpecificGenerators
{

    /// <summary>
    /// Extension of SumOtherSums built for the sole purpouse of fixing the final summary row of the ReportOutstandingBalance
    /// </summary>
    internal class FinalRowOfOutstandingBal : SumOtherSums
    {

        /// <inheritdoc/>
        public override void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {

            MoveSummaryCell(worksheet);

            base.InsertFormulas(worksheet, headers);
        }



        /// <summary>
        /// Moves the worksheet's final summary cell over 1 cell to the left. This is necessary because
        /// that cell ends up in the wrong place after merge cells are unmerged.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        private void MoveSummaryCell(ExcelWorksheet worksheet)
        {
            ExcelRange sourceCell = worksheet.Cells[worksheet.Dimension.End.Row, worksheet.Dimension.End.Column];
            ExcelRange destinationCell = worksheet.Cells[worksheet.Dimension.End.Row, worksheet.Dimension.End.Column - 1];


            destinationCell.SetCellValue(0, 0, sourceCell.Text);
            sourceCell.CopyStyles(destinationCell);
            sourceCell.SetCellValue(0, 0, "");
        }
    }
}
