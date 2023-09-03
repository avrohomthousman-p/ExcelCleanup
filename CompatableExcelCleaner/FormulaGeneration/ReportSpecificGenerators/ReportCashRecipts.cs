using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CompatableExcelCleaner.FormulaGeneration.ReportSpecificGenerators
{

    /// <summary>
    /// Version of PeriodicFormulaGenerator that is designed to work for ReportCashRecipts report
    /// </summary>
    internal class ReportCashRecipts : PeriodicFormulaGenerator
    {

        /// <inheritdoc/>
        protected override void ProcessFormulaRange(ExcelWorksheet worksheet, ref int row, int dataCol)
        {
            //TODO
        }
    }
}
