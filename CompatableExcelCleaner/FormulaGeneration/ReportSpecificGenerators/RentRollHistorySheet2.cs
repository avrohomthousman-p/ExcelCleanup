﻿using OfficeOpenXml;
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
    /// 
    /// The headers that should be passed to this class are just the names of the headers that should be split into two cells
    /// one with a total of all the cells in the monetary section of this report.
    /// one with a total of all the cells in the monetary section of this report.
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

            AddMonetarySummary(worksheet, moneySectionTop, headers);

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

            int end = iter.GetCellCoordinates(ExcelIterator.SHIFT_DOWN, cell => FormulaManager.IsEmptyCell(cell)).Last().Item1;

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

            int end = iter.GetCellCoordinates(ExcelIterator.SHIFT_DOWN, cell => FormulaManager.IsEmptyCell(cell)).Last().Item1;
            
            return new Tuple<int, int>(start, end);
        }




        /// <summary>
        /// Adds summaries to all the cells requireing them in the monetary section of the worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="startRow">the first row that is part of the monetary section</param>
        /// <param name="headers">the text marking the cells that have the formulas for the monetary section</param>
        private void AddMonetarySummary(ExcelWorksheet worksheet, int startRow, string[] headers)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, startRow, 1);
            ExcelRange startCell = iter.GetFirstMatchingCell(cell => FormulaManager.IsDollarValue(cell));
            ExcelRange endCell = iter.GetCells(ExcelIterator.SHIFT_RIGHT,
                cell => !FormulaManager.IsDollarValue(cell) && !FormulaManager.IsEmptyCell(cell)).Last();

            string formula = "SUM(" + startCell.Address + ":" + endCell.Address + ")";


            //find each summary cell and add formula to it
            foreach(string header in headers)
            {
                AddFormulaToHeader(worksheet, header, formula);
            }
        }



        /// <summary>
        /// Finds the cell that matches the specified text, splits the cell into two, and inserts the formula into the 
        /// second cell
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="header">the text of the header needing the formula</param>
        /// <param name="formula">the formula to be inserted</param>
        private void AddFormulaToHeader(ExcelWorksheet worksheet, string header, string formula)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);

            var headerCells = iter.FindAllMatchingCells(cell => FormulaManager.TextMatches(cell.Text, header));
            foreach(ExcelRange cell in headerCells)
            {
                ExcelRange formulaDestination = SplitHeaderCell(worksheet, cell);

                formulaDestination.Style.Numberformat.Format = "$#,##0.00;($#,##0.00)";
                formulaDestination.Formula = formula;
                formulaDestination.Style.Locked = true;
            }
        }



        /// <summary>
        /// Splits the specified header cell into two cells, one for the header and one for the data.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of forlumas</param>
        /// <param name="currentLocation">the cell that currently has the text</param>
        /// <returns>
        /// the cell that is now the data part of the header and needs a formula, 
        /// or null if the cell could not be split (non of the adjacent cells are availible)
        /// </returns>
        private ExcelRange SplitHeaderCell(ExcelWorksheet worksheet, ExcelRange currentLocation)
        {
            string headerText = currentLocation.Text.Substring(0, currentLocation.Text.IndexOf('$')).Trim();

            int currentRow = currentLocation.Start.Row;
            int currentCol = currentLocation.Start.Column;


            if(currentCol > 1)
            {
                ExcelRange cellToTheLeft = worksheet.Cells[currentRow, currentCol - 1];
                if (FormulaManager.IsEmptyCell(cellToTheLeft))
                {
                    //move the first half of the text over
                    cellToTheLeft.SetCellValue(0, 0, headerText);
                    currentLocation.CopyStyles(cellToTheLeft);
                    return currentLocation;
                }
            }


            ExcelRange cellToTheRight = worksheet.Cells[currentRow, currentCol + 1];
            if (FormulaManager.IsEmptyCell(cellToTheRight))
            {
                //remove the data part of the text from the current cell
                currentLocation.SetCellValue(0, 0, headerText);
                return cellToTheRight;
            }


            return null;
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
