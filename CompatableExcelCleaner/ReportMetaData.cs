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

        private static readonly string anyMonth = "(January|February|March|April|May|June|July|August|September|October|November|December)";
        private static readonly string anyDate = "\\d{2}/\\d{2}/\\d{4}";
        private static readonly string anyYear = "\\d{4}";


        // Stores the arguments needed to generate formulas for each report and worksheet. If a report/worksheet
        // is not in the dictionary, that means it doesnt need any formulas
        private static readonly Dictionary<Worksheet, string[]> formulaGenerationArguments = new Dictionary<Worksheet, string[]>();





        static ReportMetaData()
        {

            //Fill our dictionary with all the reports and all the data we need to give them formulas

            formulaGenerationArguments.Add(new Worksheet("ProfitAndLossStatementByPeriod", 0), new String[] { "Total Income", "^Total Expense" });
            formulaGenerationArguments.Add(new Worksheet("LedgerReport", 0), new String[] { "Total \\d+ - Prepaid Contracts" }); //ISSUE: numbers dont add up
            formulaGenerationArguments.Add(new Worksheet("RentRollAll", 0), new String[] { "Total:" });
            formulaGenerationArguments.Add(new Worksheet("ProfitAndLossStatementDrillthrough", 0), new String[] { "Total Expense", "Total Income" });
            formulaGenerationArguments.Add(new Worksheet("ProfitAndLossStatementDrillthrough", 1), new String[] { "Total Expense", "Total Income" });
            formulaGenerationArguments.Add(new Worksheet("BalanceSheetDrillthrough", 0), new String[]
                    { "Current Assets=Total Current Assets", "Fixed Asset=Total Fixed Asset", "Other Asset=Total Other Asset",
                        "Assets=Total Assets", "Liabilities And Equity=Total Liabilities And Equity",
                        "Current Liabilities=Total Current Liabilities", "Liability=Total Liability",
                        "Long Term Liability=Total Long Term Liability", "Equity=Total Equity",
                        "Total Liabilities~Total Long Term Liability,Total Liability,Total Current Liabilities"
                    });

            formulaGenerationArguments.Add(new Worksheet("ReportTenantBal", 0), new String[] { "Total Open Charges:", "Balance:~Total Open Charges:,Total Future Charges:,Total Unallocated Payments:" });
            //ISSUE last formula missing
            formulaGenerationArguments.Add(new Worksheet("ReportOutstandingBalance", 0), new String[] { "Balance" });
            formulaGenerationArguments.Add(new Worksheet("ReportOutstandingBalance", 1), new String[] { "Total" });
            formulaGenerationArguments.Add(new Worksheet("BalanceSheetComp", 0), new String[] 
            { "Current Assets=Total Current Assets", "Fixed Asset=Total Fixed Asset", "Other Asset=Total Other Asset", 
                "Assets=Total Assets",  "Current Liabilities=Total Current Liabilities", "Liability=Total Liability",
                "Liabilities And Equity=Total Liabilities And Equity", "Long Term Liability=Total Long Term Liability",
                "Equity=Total Equity"}); //ISSUE: one header skipped

            formulaGenerationArguments.Add(new Worksheet("AgedReceivables", 0), new String[] { "Total" });
            formulaGenerationArguments.Add(new Worksheet("ProfitAndLossComp", 0), new String[] { "INCOME=Total Income", "EXPENSE=Total Expense" });
            formulaGenerationArguments.Add(new Worksheet("RentRollActivity_New", 0), new String[] { "Total:" });
            formulaGenerationArguments.Add(new Worksheet("RentRollActivity_New", 1), new String[] { "Total For International City:" });
            formulaGenerationArguments.Add(new Worksheet("TrialBalance", 0), new String[] { "Total:" });



            formulaGenerationArguments.Add(new Worksheet("ReportCashReceiptsSummary", 0), new String[] {
                        "Total Tenant Receivables:", "Total Other Receivables:",
                        $"Total For {anyMonth} {anyYear}:~Total Tenant Receivables:,Total Other Receivables:",
                        $"Total For Commons at White Marsh:~Total For {anyMonth} {anyYear}:"});

            formulaGenerationArguments.Add(new Worksheet("ReportCashReceiptsSummary", 1), new String[] { });


            formulaGenerationArguments.Add(new Worksheet("ReportPayablesRegister", 0), new String[] { }); //corrupted file
            formulaGenerationArguments.Add(new Worksheet("AgedPayables", 0), new String[] { "Total" });
            formulaGenerationArguments.Add(new Worksheet("ChargesCreditReport", 0), new String[] { "Total: $(\\d\\d\\d,)*\\d?\\d?\\d[.]\\d\\d" });//fixme: needs special system
            formulaGenerationArguments.Add(new Worksheet("UnitInvoiceReport", 0), new String[] { });//dont have this report
            formulaGenerationArguments.Add(new Worksheet("ReportCashReceipts", 0), new String[] { }); //needs new system
            formulaGenerationArguments.Add(new Worksheet("PayablesAccountReport", 0), new String[] { 
                "Pool Furniture=Total Pool Furniture", "Hallways=Total Hallways", "Garage=Total Garage", 
                "Elevators=Total Elevators", "Clubhouse=Total Clubhouse", 
                "Total Common Area CapEx~Total Pool Furniture,Total Hallways,Total Garage,Total Elevators,Total Clubhouse", 
                "Total~Total Common Area CapEx", "Total:~Total Common Area CapEx" });






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
                case "ReportCashReceiptsSummary":
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
        /// <returns>
        /// an implemenation of the IFormulaGenerator interface that should be used to add the formulas,
        /// or null if the worksheet doesnt need formulas
        /// </returns>
        internal static IFormulaGenerator ChooseFormulaGenerator(string reportName, int worksheetNum)
        {
            switch (reportName)
            {
                case "BalanceSheetDrillthrough":
                case "BalanceSheetComp":
                case "ProfitAndLossComp":
                case "PayablesAccountReport":
                    return new RowSegmentFormulaGenerator();



                case "ReportOutstandingBalance":
                    switch (worksheetNum)
                    {
                        case 0:
                            return new DataColumnFormulaGenerator();
                        default:
                            return new FullTableFormulaGenerator();
                    }


                
                case "AgedPayables":
                case "AgedReceivables":
                    return new FullTableFormulaGenerator(FullTableFormulaGenerator.IsNonDataCell);



                case "ReportTenantBal":
                case "ProfitAndLossStatementByPeriod":
                case "LedgerReport":
                case "RentRollAll":
                case "ProfitAndLossStatementDrillthrough": 
                case "RentRollActivity_New":
                case "TrialBalance":
                case "ReportCashReceiptsSummary":
                    return new FullTableFormulaGenerator();




                //These reports dont fit into any existing system
                case "ChargesCreditReport":
                case "ReportCashReceipts":





                //Reports I dont have
                case "ReportPayablesRegister":
                case "UnitInvoiceReport":




                //Reports I have not yet checked
                case "VendorInvoiceReport":
                case "CollectionsAnaysisSummary":
                case "ReportTenantSummary":
                case "ProfitAndLossBudget":
                case "VendorInvoiceReportWithJournalAccounts":
                case "RentRollActivityItemized_New":
                case "RentHistoryReport":
                case "ProfitAndLossExtendedVariance":
                case "RentRollActivity":
                case "RentRollAllItemized":
                case "RentRollHistory":
                case "TrialBalanceVariance":
                case "JournalLedger":
                case "CollectionsAnalysis":
                case "ProfitAndLossStatementByJob":
                case "VendorPropertyReport":
                case "RentRollPortfolio":
                case "AgedAccountsReceivable":
                case "BalanceSheetPropBreakdown":
                case "SubsidyRentRollReport":
                case "VacancyLoss":
                case "PropBankAccountReport":
                case "PayablesAuditTrail":
                case "PaymentsHistory":
                case "MarketRentReport":
                case "Budget":
                case "ReportAccountBalances":
                case "CCTransactionsReport":
                case "RentRollActivityCompSummary":
                case "RentRollCommercialItemized":
                case "RentRollActivityTotals":
                case "RentRollBalanceHistory":
                case "PreprintedLeasesReport":
                case "ReportEscalateCharges":
                case "RentRollActivityItemized":
                case "InvoiceRecurringReport":
                    


                default:
                    return null;
            }
        }



        /// <summary>
        /// Retrieves the required arguments that should be passed into IFormulaGenerator.InsertFormulas function
        /// for a given report and worksheet.
        /// </summary>
        /// <param name="reportName">the name of the report getting the formulas</param>
        /// <param name="worksheetNum">the index of the worksheet getting the formulas</param>
        /// <returns>
        /// a list of strings that should be passed to the formula generator when formulas are being added,
        /// or null if the worksheet does not require formulas
        /// </returns>
        internal static string[] GetFormulaGenerationArguments(string reportName, int worksheetNum)
        {
            return formulaGenerationArguments[  new Worksheet(reportName, worksheetNum)  ];
        }



        /// <summary>
        /// Checks if the specified worksheet in the specified report requires formulas
        /// </summary>
        /// <param name="reportName">the name of the report</param>
        /// <param name="worksheetNumber">the index of the worksheet</param>
        /// <returns>true if the worksheet specified needs formulas and false otherwise</returns>
        internal static bool RequiresFormulas(string reportName, int worksheetNumber)
        {
            string[] ignoredResult = new string[0];

            return formulaGenerationArguments.TryGetValue(new Worksheet(reportName, worksheetNumber), out ignoredResult);
        }
    }
}
