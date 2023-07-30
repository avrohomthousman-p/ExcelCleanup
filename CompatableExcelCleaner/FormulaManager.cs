using OfficeOpenXml;
using System;
using System.IO;
using System.Collections.Generic;


namespace CompatableExcelCleaner
{
    /// <summary>
    /// Replaces static values in excel files with formulas that will change when the data is updated
    /// </summary>
    public class FormulaManager
    {


        // Stores each report that requires formulas and an array of all the headers to look for in that report
        // (for where to add those formulas)
        private static readonly Dictionary<string, string[]> rowsNeedingFormulas = new Dictionary<string, string[]>();



        static FormulaManager() 
        {

            //Fill our dictionary with all the reports and all the data we need to give them formulas

            rowsNeedingFormulas.Add("ProfitAndLossStatementByPeriod", new String[]{ "Total Income", "Total Expense" });
            rowsNeedingFormulas.Add("LedgerReport", new String[] { "14850 - Prepaid Contracts" }); //ISSUE
            rowsNeedingFormulas.Add("RentRollAll", new String[] { "Total:" });
            rowsNeedingFormulas.Add("ProfitAndLossStatementDrillthrough", new String[] { "Total Expense", "Total Income" }); //ISSUE

            //TODO: tryout all these reports
            rowsNeedingFormulas.Add("BalanceSheetDrillThrough", new String[] { });
            rowsNeedingFormulas.Add("ReportTenantBal", new String[] { });
            rowsNeedingFormulas.Add("ReportOutstandingBalance", new String[] { });
            rowsNeedingFormulas.Add("BalanceSheetComp", new String[] { });
            rowsNeedingFormulas.Add("AgedReceivables", new String[] { });
            rowsNeedingFormulas.Add("ProfitAndLossComp", new String[] { });
            rowsNeedingFormulas.Add("RentRollActivity_New", new String[] { });
            rowsNeedingFormulas.Add("TrialBalance", new String[] { });
            rowsNeedingFormulas.Add("ReportCashReceiptsSummary", new String[] { });
            rowsNeedingFormulas.Add("ReportPayablesRegister", new String[] { });
            rowsNeedingFormulas.Add("AgedPayables", new String[] { });
            rowsNeedingFormulas.Add("ChargesCreditReport", new String[] { });
            rowsNeedingFormulas.Add("UnitInvoiceReport", new String[] { });
            rowsNeedingFormulas.Add("ReportCashReceipts", new String[] { });
            rowsNeedingFormulas.Add("PayablesAccountReport", new String[] { });
            rowsNeedingFormulas.Add("VendorInvoiceReport", new String[] { });
            rowsNeedingFormulas.Add("CollectionsAnaysisSummary", new String[] { });
            rowsNeedingFormulas.Add("ReportTenantSummary", new String[] { });
            rowsNeedingFormulas.Add("ProfitAndLossBudget", new String[] { });
            rowsNeedingFormulas.Add("VendorInvoiceReportWithJournalAccounts", new String[] { });
            rowsNeedingFormulas.Add("RentRollActivityItemized_New", new String[] { });
            rowsNeedingFormulas.Add("RentHistoryReport", new String[] { });
            rowsNeedingFormulas.Add("ProfitAndLossExtendedVariance", new String[] { });
            rowsNeedingFormulas.Add("RentRollActivity", new String[] { });
            rowsNeedingFormulas.Add("RentRollAllItemized", new String[] { });
            rowsNeedingFormulas.Add("RentRollHistory", new String[] { });
            rowsNeedingFormulas.Add("TrialBalanceVariance", new String[] { });
            rowsNeedingFormulas.Add("JournalLedger", new String[] { });
            rowsNeedingFormulas.Add("CollectionsAnalysis", new String[] { });
            rowsNeedingFormulas.Add("ProfitAndLossStatementByJob", new String[] { });
            rowsNeedingFormulas.Add("VendorPropertyReport", new String[] { });
            rowsNeedingFormulas.Add("RentRollPortfolio", new String[] { });
            rowsNeedingFormulas.Add("AgedAccountsReceivable", new String[] { });
            rowsNeedingFormulas.Add("BalanceSheetPropBreakdown", new String[] { });
            rowsNeedingFormulas.Add("SubsidyRentRollReport", new String[] { });
            rowsNeedingFormulas.Add("VacancyLoss", new String[] { });
            rowsNeedingFormulas.Add("PropBankAccountReport", new String[] { });
            rowsNeedingFormulas.Add("PayablesAuditTrail", new String[] { });
            rowsNeedingFormulas.Add("PaymentsHistory", new String[] { });
            rowsNeedingFormulas.Add("MarketRentReport", new String[] { });
            rowsNeedingFormulas.Add("Budget", new String[] { });
            rowsNeedingFormulas.Add("ReportAccountBalances", new String[] { });
            rowsNeedingFormulas.Add("CCTransactionsReport", new String[] { });
            rowsNeedingFormulas.Add("RentRollActivityCompSummary", new String[] { });
            rowsNeedingFormulas.Add("RentRollCommercialItemized", new String[] { });
            rowsNeedingFormulas.Add("RentRollActivityTotals", new String[] { });
            rowsNeedingFormulas.Add("RentRollBalanceHistory", new String[] { });
            rowsNeedingFormulas.Add("PreprintedLeasesReport", new String[] { });
            rowsNeedingFormulas.Add("ReportEscalateCharges", new String[] { });
            rowsNeedingFormulas.Add("RentRollActivityItemized", new String[] { });
            rowsNeedingFormulas.Add("InvoiceRecurringReport", new String[] { });
        }



        /// <summary>
        /// Adds all necissary formulas to the appropriate cells in the specified file
        /// </summary>
        /// <param name="sourceFile">the excel file needing formulas, stored as an array/stream of bytes</param>
        /// <param name="reportName">the name of the report</param>
        /// <returns>the byte stream/arrray of the modified file</returns>
        public static byte[] AddFormulas(byte[] sourceFile, string reportName)
        {
            string[] headers;
            if(!rowsNeedingFormulas.TryGetValue(reportName, out headers))
            {
                return sourceFile; //this report is not supposed to get any formulas
            }


            using (ExcelPackage package = new ExcelPackage(new MemoryStream(sourceFile)))
            {
                ExcelWorksheet worksheet;
                for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
                {
                    worksheet = package.Workbook.Worksheets[i];

                    //call formula generator
                    IFormulaGenerator formulaGenerator = ChooseFormulaGenerator(reportName);
                    formulaGenerator.InsertFormulas(worksheet, headers);
                }


                return package.GetAsByteArray();
            }

        }



        /// <summary>
        /// Chooses the implementation of the IFormulaGenerator interface that should be used to add formulas
        /// to the specified report.
        /// </summary>
        /// <param name="reportName">the name of the report that needs formulas</param>
        /// <returns>an implemenation of the IFormulaGenerator interface that should be used to add the formulas</returns>
        private static IFormulaGenerator ChooseFormulaGenerator(string reportName)
        {
            switch(reportName)
            {
                case "TODO":
                    return new RowSegmentFormulaGenerator();


                //case "ProfitAndLossDrillthrough":
                //case "ProfitAndLossStatementByPeriod":
                //case "RentRollAll":
                default:
                    return new FullTableFormulaGenerator();
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
        /// Checks if a cell contains a dollar value
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell contains a dollar value and false otherwise</returns>
        internal static bool IsDataCell(ExcelRange cell)
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
            //ExcelRange cells = worksheet.Cells[startRow + 1, col, endRow - 1, col];
            ExcelRange cells = worksheet.Cells[startRow, col, endRow, col];

            return "SUM(" + cells.Address + ")";
        }

    }
}
