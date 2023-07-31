using ExcelDataCleanup;
using System;
using System.Collections.Generic;

namespace CompatableExcelCleaner
{

    /// <summary>
    /// Stores and makes accesible all meta data about reports, like what merge cleaner and formula generator to
    /// use.
    /// </summary>
    internal static class ReportMetaData
    {


        // Stores the arguments needed to generate formulas for each report and worksheet. If a report/worksheet
        // is not in the dictionary, that means it doesnt need any formulas
        private static readonly Dictionary<Worksheet, string[]> formulaGenerationArguments = new Dictionary<Worksheet, string[]>();




        static ReportMetaData()
        {

            //Fill our dictionary with all the reports and all the data we need to give them formulas

            formulaGenerationArguments.Add(new Worksheet("ProfitAndLossStatementByPeriod", 0), new String[] { "Total Income", "Total Expense" });
            formulaGenerationArguments.Add(new Worksheet("LedgerReport", 0), new String[] { "14850 - Prepaid Contracts" }); //ISSUE: numbers dont add up
            formulaGenerationArguments.Add(new Worksheet("RentRollAll", 0), new String[] { "Total:" });
            formulaGenerationArguments.Add(new Worksheet("ProfitAndLossStatementDrillthrough", 0), new String[] { "Total Expense", "Total Income" });
            formulaGenerationArguments.Add(new Worksheet("BalanceSheetDrillthrough", 0), new String[]
                    { "Current Assets=Total Current Assets", "Fixed Asset=Total Fixed Asset", "Other Asset=Total Other Asset",
                        "Assets=Total Assets", "Liabilities And Equity=Total Liabilities And Equity",
                        "Current Liabilities=Total Current Liabilities", "Liability=Total Liability",
                        "Long Term Liability=Total Long Term Liability", "Equity=Total Equity" }); //SMALL ISSUE: one line isnt getting formula

            //ISSUE: small empty rows that have not been deleted
            formulaGenerationArguments.Add(new Worksheet("ReportTenantBal", 0), new String[] { "Total Open Charges:=Balance:", "Electric Bill: 08/25/2022-09/28/2022=Trash" });




            //TODO: tryout all these reports
            formulaGenerationArguments.Add(new Worksheet("ReportOutstandingBalance", 0), new String[] { });


            formulaGenerationArguments.Add(new Worksheet("BalanceSheetComp", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("AgedReceivables", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("ProfitAndLossComp", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("RentRollActivity_New", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("TrialBalance", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("ReportCashReceiptsSummary", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("ReportPayablesRegister", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("AgedPayables", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("ChargesCreditReport", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("UnitInvoiceReport", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("ReportCashReceipts", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("PayablesAccountReport", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("VendorInvoiceReport", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("CollectionsAnaysisSummary", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("ReportTenantSummary", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("ProfitAndLossBudget", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("VendorInvoiceReportWithJournalAccounts", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("RentRollActivityItemized_New", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("RentHistoryReport", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("ProfitAndLossExtendedVariance", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("RentRollActivity", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("RentRollAllItemized", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("RentRollHistory", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("TrialBalanceVariance", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("JournalLedger", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("CollectionsAnalysis", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("ProfitAndLossStatementByJob", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("VendorPropertyReport", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("RentRollPortfolio", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("AgedAccountsReceivable", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("BalanceSheetPropBreakdown", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("SubsidyRentRollReport", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("VacancyLoss", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("PropBankAccountReport", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("PayablesAuditTrail", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("PaymentsHistory", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("MarketRentReport", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("Budget", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("ReportAccountBalances", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("CCTransactionsReport", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("RentRollActivityCompSummary", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("RentRollCommercialItemized", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("RentRollActivityTotals", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("RentRollBalanceHistory", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("PreprintedLeasesReport", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("ReportEscalateCharges", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("RentRollActivityItemized", 0), new String[] { });
            formulaGenerationArguments.Add(new Worksheet("InvoiceRecurringReport", 0), new String[] { });
        }




        /// <summary>
        /// Factory method for choosing a version of the merge cleanup code that would work best for the specified report
        /// </summary>
        /// <param name="reportType">the type of report that needs unmerging</param>
        /// <param name="worksheetNumber">the worksheet withing the report that needs unmerging</param>
        /// <returns>an instance of IMergeCleaner that should be used to clean the specified worksheet</returns>
        internal static IMergeCleaner ChoosesCleanupSystem(string reportType, int worksheetNumber)
        {
            switch (reportType)
            {
                case "TrialBalance":
                case "TrialBalanceVariance":
                case "ProfitAndLossStatementDrillthrough":
                case "BalanceSheetDrillthrough":
                case "CashFlow":
                case "InvoiceDetail":
                case "ReportTenantSummary":
                case "UnitInfoReport":
                    return new BackupMergeCleaner();



                case "ReportOutstandingBalance":
                    switch (worksheetNumber)
                    {
                        case 1:
                            return new BackupMergeCleaner();
                        default:
                            return new PrimaryMergeCleaner();
                    }



                default:
                    return new PrimaryMergeCleaner();
            }
        }




        /// <summary>
        /// Factory method for choosing the implementation of the IFormulaGenerator interface that should be used to add formulas
        /// to the specified report.
        /// </summary>
        /// <param name="reportName">the name of the report that needs formulas</param>
        /// <param name="worksheetNum">the index of the worksheet that needs formulas</param>
        /// <returns>an implemenation of the IFormulaGenerator interface that should be used to add the formulas</returns>
        internal static IFormulaGenerator ChooseFormulaGenerator(string reportName, int worksheetNum)
        {
            switch (reportName)
            {
                case "BalanceSheetDrillthrough":
                case "ReportTenantBal":
                    return new RowSegmentFormulaGenerator();

                default:
                    return new FullTableFormulaGenerator();
            }
        }



        /// <summary>
        /// Retrieves the required arguments that should be passed into IFormulaGenerator.InsertFormulas function
        /// for a given report and worksheet.
        /// </summary>
        /// <param name="reportName">the name of the report getting the formulas</param>
        /// <param name="worksheetNum">the index of the worksheet getting the formulas</param>
        /// <returns>a list of strings that should be passed to the formula generator when formulas are being added</returns>
        internal static string[] GetFormulaGenerationArguments(string reportName, int worksheetNum)
        {
            return formulaGenerationArguments[  new Worksheet(reportName, worksheetNum)  ];
        }
    }
}
