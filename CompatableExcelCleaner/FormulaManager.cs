using OfficeOpenXml;
using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace CompatableExcelCleaner
{
    /// <summary>
    /// Replaces static values in excel files with formulas that will change when the data is updated
    /// </summary>
    public class FormulaManager
    {



        /// <summary>
        /// Adds all necissary formulas to the appropriate cells in the specified file
        /// </summary>
        /// <param name="sourceFile">the excel file needing formulas, stored as an array/stream of bytes</param>
        /// <param name="reportName">the name of the report</param>
        /// <returns>the byte stream/arrray of the modified file</returns>
        public static byte[] AddFormulas(byte[] sourceFile, string reportName)
        {
            
            using (ExcelPackage package = new ExcelPackage(new MemoryStream(sourceFile)))
            {

                string[] headers;
                ExcelWorksheet worksheet;

                for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
                {
                    worksheet = package.Workbook.Worksheets[i];



                    //If the worksheet is empty, Dimension will be null
                    if (worksheet.Dimension == null)
                    {
                        package.Workbook.Worksheets.Delete(i);
                        i--;
                        continue;
                    }


                    IFormulaGenerator formulaGenerator = ReportMetaData.ChooseFormulaGenerator(reportName, i);

                    if(formulaGenerator == null) //if this worksheet doesnt need formulas
                    {
                        continue; //skip this worksheet
                    }

                    headers = ReportMetaData.GetFormulaGenerationArguments(reportName, i);

                    formulaGenerator.InsertFormulas(worksheet, headers);


                    //Add formulas for rows that are not contiguous if needed
                    SummaryRowFormulaGenerator.InsertFormulas(worksheet, headers);
                }


                return package.GetAsByteArray();
            }

        }





        /* Some utility methods needed by the Formula generators */


        /// <summary>
        /// Checks if a cell is empty (has no text)
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell has no text and false otherwise</returns>
        internal static bool IsEmptyCell(ExcelRange cell)
        {
            return cell.Text == null || cell.Text.Length == 0;
        }



        /// <summary>
        /// Checks if a cell contains a dollar value. This is used as a default implementation for IsDataCell in 
        /// formula managers. Formula managers can be set to use a different defenition for a data cell 
        /// by calling the method IFormulaManager.SetDataCellDefenition(  specify alternate implementation  )
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell contains a dollar value and false otherwise</returns>
        internal static bool IsDollarValue(ExcelRange cell)
        {
            return cell.Text.StartsWith("$") || (cell.Text.StartsWith("($") && cell.Text.EndsWith(")"));
        }



        /// <summary>
        /// Generates the formula for the cells in the given range. Note: the range should only include the 
        /// cells that are to be included in the formula. Not the that cell that will contain the formula itself
        /// or any cells above the range.
        /// </summary>
        /// <param name="worksheet">the worksheet currently getting formulas</param>
        /// <param name="startRow">the first data cell to be included in the formula</param>
        /// <param name="endRow">the last data cell to be included in the formula</param>
        /// <param name="col">the column the formula is for</param>
        /// <returns>the proper formula for the specified formula range</returns>
        internal static string GenerateFormula(ExcelWorksheet worksheet, int startRow, int endRow, int col)
        {
            ExcelRange cells = worksheet.Cells[startRow, col, endRow, col];

            return "SUM(" + cells.Address + ")";
        }




        /// <summary>
        /// Checks if a header is intened for the DistantRowsFormulaGenerator or not.
        /// </summary>
        /// <param name="header">the header in question</param>
        /// <returns>true if the specified header is intended for the DistantRowsFormulaGenerator class, and false otherwise</returns>
        internal static bool IsNonContiguousFormulaRange(string header)
        {
            return header.IndexOf('~') >= 0;
        }


        

        /// <summary>
        /// Checks if the specified text matches (in its entirety) the specified regex.
        /// </summary>
        /// <param name="text">the text to be matched</param>
        /// <param name="pattern">the pattern the text should match</param>
        /// <returns>true if the text matches the pattern and false otherwise</returns>
        internal static bool TextMatches(string text, string pattern)
        {
            return Regex.IsMatch(text.Trim(), "^" + pattern + "$");
        }

    }


}
