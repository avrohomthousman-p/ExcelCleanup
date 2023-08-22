using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace CompatableExcelCleaner.FormulaGeneration.ReportSpecificGenerators
{


    /// <summary>
    /// Implementation of IFormulaGenerator that is specifically designed to add formulas to the second sheet 
    /// of the RentRollHistory report, which can't really work with any existing system.
    /// </summary>
    internal class RentRollHistorySheet2 : IFormulaGenerator
    {
        private static readonly string MONTH_REGEX = "(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) (19|20)\\d\\d";

        private Predicate<ExcelRange> isMonth = cell => FormulaManager.TextMatches(cell.Text, MONTH_REGEX);

        private IsDataCell dataCellDef = new IsDataCell(
                                    cell => FormulaManager.IsDollarValue(cell) 
                                        || FormulaManager.IsIntegerValue(cell) 
                                        || FormulaManager.IsPercentage(cell));


        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {

            Tuple<int, int> rowRange = FindMoneySection(worksheet);
            int moneySectionTop = rowRange.Item1;
            int moneySectionBottom = rowRange.Item2;

            rowRange = FindOccupancySection(worksheet, moneySectionBottom - 1);
            int occupancySectionTop = rowRange.Item1;
            int occupancySectionBottom = rowRange.Item2;


            FormattingCleanup(worksheet);

            AddMonetarySummary(worksheet, moneySectionTop, moneySectionBottom);

            AddOccupancySummaries(worksheet, occupancySectionTop, occupancySectionBottom);
        }



        /// <summary>
        /// Finds the start and end rows of the section of the worksheet that has financial data in it
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <returns>a tuple with the start and end rows of the money section</returns>
        private Tuple<int, int> FindMoneySection(ExcelWorksheet worksheet)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);

            int start = iter.GetFirstMatchingCell(isMonth).Start.Row;

            int end = iter.GetCellCoordinates(ExcelIterator.SHIFT_DOWN, cell => FormulaManager.IsEmptyCell(cell)).Last().Item1 - 1;
            //SAFE TO MAKE CHANGE HERE
            return new Tuple<int, int>(start, end);
        }




        /// <summary>
        /// Finds the start and end rows of the section of the worksheet that has data about occupancy in it
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="startSearchAtRow">the row to start searching at, which should be just after the money section</param>
        /// <returns>a tuple with the start and end rows of the occupancy section</returns>
        private Tuple<int, int> FindOccupancySection(ExcelWorksheet worksheet, int startSearchAtRow)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, startSearchAtRow, 1);

            int start = iter.GetFirstMatchingCell(isMonth).Start.Row;

            int end = iter.GetCellCoordinates(ExcelIterator.SHIFT_DOWN, cell => FormulaManager.IsEmptyCell(cell)).Last().Item1 - 1;
            //SAFE TO MAKE CHANGE HERE
            return new Tuple<int, int>(start, end);
        }




        /// <summary>
        /// Adds summaries to all the cells requireing them in the monetary section of the worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="startRow">the first row that is part of the monetary section</param>
        /// <param name="endRow">the last row that is part of the monetary section</param>
        private void AddMonetarySummary(ExcelWorksheet worksheet, int startRow, int endRow)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, startRow, 1);
            ExcelRange startCell = iter.GetFirstMatchingCell(cell => FormulaManager.IsDollarValue(cell));
            ExcelRange endCell = iter.GetCells(ExcelIterator.SHIFT_RIGHT, //SAFE TO MAKE CHANGE
                cell => !FormulaManager.IsDollarValue(cell) && !FormulaManager.IsEmptyCell(cell)).Last();

            //TODO
        }



        /// <summary>
        /// Adds summaries to all the cells requireing them in the occupancy section of the worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="startRow">the first row that is part of the occupancy section</param>
        /// <param name="endRow">the last row that is part of the occupancy section</param>
        private void AddOccupancySummaries(ExcelWorksheet worksheet, int startRow, int endRow)
        {
            //TODO
        }




        /// <summary>
        /// The RentRollHistory report has some formatting issues that should be fixed
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        private void FormattingCleanup(ExcelWorksheet worksheet)
        {
            //TODO
        }




        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.dataCellDef = isDataCell;
        }
    }
}
