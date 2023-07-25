using OfficeOpenXml;
using System;
using System.IO;


namespace CompatableExcelCleaner
{
    /// <summary>
    /// Replaces static values in excel files with formulas that will change when the data is updated
    /// </summary>
    public class FormulaMaker
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
                ExcelWorksheet worksheet;
                for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
                {
                    worksheet = package.Workbook.Worksheets[i];
                    
                    //TODO: add formulas

                }
            }

            return sourceFile;
        }



        /// <summary>
        /// Gets the row numbers of the first and last rows that should be included in the formula
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="targetText">the text to look for to signal the start and end row</param>
        /// <returns>a tuple containing the start-row and end-row of the formula range</returns>
        private static Tuple<int, int> GetRowRangeForFormula(ExcelWorksheet worksheet, string targetText)
        {
            ExcelRange cell;


            for(int row = 1; row < worksheet.Dimension.End.Row; row++)
            {
                for(int col = 1; col < worksheet.Dimension.End.Column; col++)
                {
                    cell = worksheet.Cells[row, col];

                    if (cell.Text == targetText)
                    {
                        //TODO: search for end of sequence
                    }
                }
            }

            return null; //FIXME
        }




        /// <summary>
        /// Given the Cell coordinates of the starting cell in a formula range, finds the ending cell for 
        /// that range.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being given formulas</param>
        /// <param name="row">the row number of the starting cell in the formula range</param>
        /// <param name="col">the column number of the starting cell in the formula range</param>
        /// <param name="targetText">the text to look for that signals the end cell of the formula range</param>
        /// <returns>the row number of the last cell in the formula range, or -1 if no appropriate last cell is found</returns>
        private static int FindEndOfFormulaRange(ExcelWorksheet worksheet, int row, int col, string targetText)
        {
            ExcelRange cell;

            for(int i = row + 1; i < worksheet.Dimension.End.Row; i++)
            {
                cell = worksheet.Cells[i, col];
                if (cell.Text == targetText)
                {
                    return i;
                }
            }


            return -1;
        }
    }
}
