using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompatableExcelCleaner
{
    internal interface IFormulaGenerator
    {
        /// <summary>
        /// Finds each instance of each header in the worksheet and gives its corisponding columns a formula.
        /// The term "header" here refers to minor headers in the worksheet with text like "Total Income" that 
        /// indicate that the row they are in is a summary row whose value should be calculated by aformula and 
        /// not just be a static value.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="headers">an array of all minor headers that are right in front of cells requiring formulas</param>
        void InsertFormulas(ExcelWorksheet worksheet, string[] headers);
    }
}
