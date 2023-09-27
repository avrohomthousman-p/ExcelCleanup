﻿using CompatableExcelCleaner.FormulaGeneration;
using CompatableExcelCleaner.FormulaGeneration.ReportSpecificGenerators;
using CompatableExcelCleaner.GeneralCleaning;
using ExcelDataCleanup;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.Xml;
using System.Text.RegularExpressions;

namespace CompatableExcelCleaner
{

    /// <summary>
    /// Stores and makes accesible all meta data about reports, like what merge cleaner and formula generator to
    /// use.
    /// </summary>
    internal static class ReportMetaData
    {

        private static readonly string anyMonth = "(January|February|March|April|May|June|July|August|September|October|November|December)";
        private static readonly string anyDate = "\\d{1,2}/\\d{1,2}/\\d{4}";
        private static readonly string anyYear = "[12]\\d\\d\\d";




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
                case "ProfitAndLossStatementDrillThrough":
                case "BalanceSheetDrillthrough":
                case "CashFlow":
                case "InvoiceDetail":
                case "ReportTenantSummary":
                case "UnitInfoReport":
                case "ReportCashReceiptsSummary":
                    return new BackupMergeCleaner();



                case "ProfitAndLossExtendedVariance":
                    return new ExtendedVarianceCleaner();



                case "ReportOutstandingBalance":
                    switch (worksheetNumber)
                    {
                        case 1:
                            return new BackupMergeCleaner();
                        default:
                            return new PrimaryMergeCleaner();
                    }



                case "RentRollHistory":
                    switch (worksheetNumber)
                    {
                        case 1:
                            return new ReAlignDataCells("Vacancy %");
                        default:
                            return new PrimaryMergeCleaner();
                    }



                case "Budget":
                    return new ReAlignDataCells();



                default:
                    return new PrimaryMergeCleaner();
            }
        }



        /// <summary>
        /// Checks if the specified worksheet needs to have its summary cells shifted one cell to the left.
        /// Due to a bug in the report generator, some reports have their summary cells one cell too far
        /// to the right.
        /// </summary>
        /// <param name="reportName">the name of the report the worksheet is from</param>
        /// <param name="worksheetIndex">the zero based index of the worksheet</param>
        /// <returns>true if the worksheets needs its summary cells moved, and false otherwise</returns>
        internal static bool NeedsSummaryCellsMoved(string reportName, int worksheetIndex)
        {
            switch (reportName)
            {
                //TODO: add the other reports with this issue
                case "ReportOutstandingBalance":
                case "ProfitAndLossBudget":
                    return true;


                default:
                    return false;
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

            FullTableFormulaGenerator formulaGenerator;


            switch (reportName)
            {
                case "BalanceSheetDrillthrough":
                case "BalanceSheetComp":
                case "ProfitAndLossComp":
                case "PayablesAccountReport":
                case "ProfitAndLossBudget":
                case "BalanceSheetPropBreakdown":
                case "ProfitAndLossStatementDrillThrough":
                case "ProfitAndLossExtendedVariance":
                    return new RowSegmentFormulaGenerator();



                case "TrialBalanceVariance":
                case "ProfitAndLossStatementByJob":
                    RowSegmentFormulaGenerator gen = new RowSegmentFormulaGenerator();
                    gen.trimFormulaRange = false;
                    return gen;




                case "ReportOutstandingBalance":
                    switch (worksheetNum)
                    {
                        case 0:
                            return new MultiFormulaGenerator(new PeriodicFormulaGenerator(), new SumOtherSums());
                        default:
                            return new FullTableFormulaGenerator();
                    }



                case "RentRollActivityItemized":
                case "RentRollActivityItemized_New":
                    PeriodicFormulaGenerator mainFormulas = new PeriodicFormulaGenerator();
                    mainFormulas.SetDataCellDefenition(cell => FormulaManager.IsEmptyCell(cell) || FormulaManager.IsDollarValue(cell));

                    SumOtherSums otherFormulas = new SumOtherSums();

                    return new MultiFormulaGenerator(mainFormulas, otherFormulas);



                case "RentRollHistory":
                    switch (worksheetNum)
                    {
                        case 0:
                            return new RentRollHistorySheet1();
                        case 1:
                            return new RentRollHistorySheet2();
                        default:
                            return null;
                    }



                case "VendorInvoiceReportWithJournalAccounts":
                    switch (worksheetNum)
                    {
                        case 5:
                            return new FullTableFormulaGenerator();
                        default:
                            return new VendorInvoiceReportFormulas();
                    }



                case "RentRollAllItemized":
                    switch (worksheetNum)
                    {
                        case 2:
                            FullTableFormulaGenerator first = new FullTableFormulaGenerator();
                            RowSegmentFormulaGenerator second = new RowSegmentFormulaGenerator();
                            IsDataCell dataCellDef = new IsDataCell(cell => 
                                    FormulaManager.IsDollarValue(cell) 
                                    || FormulaManager.IsIntegerWithCommas(cell) 
                                    || FormulaManager.IsPercentage(cell)
                                    || FormulaManager.CellHasFormula(cell));


                            first.SetDefenitionForBeyondFormulaRange(first.IsNonDataCell);

                            MultiFormulaGenerator generator = new MultiFormulaGenerator(first, second);
                            generator.SetDataCellDefenition(dataCellDef);
                            return generator;


                        case 1:
                            return new MultiFormulaGenerator(new PeriodicFormulaGenerator(), new SumOtherSums(), new FormulaBetweenSheets());
                        default:
                            return new MultiFormulaGenerator(new PeriodicFormulaGenerator(), new SumOtherSums());
                    }



                case "VacancyLoss":
                    switch (worksheetNum)
                    {
                        case 0:
                            formulaGenerator = new FullTableFormulaGenerator();
                            int ignoredOutput;
                            formulaGenerator.SetDataCellDefenition(cell => FormulaManager.IsDollarValue(cell) || Int32.TryParse(cell.Text, out ignoredOutput));
                            return formulaGenerator;

                        default:
                            return new FullTableFormulaGenerator();
                    }



                case "ReportCashReceipts":
                    return new ReportCashRecipts();



                case "ChargesCreditsReport":
                    return new ChargesCreditReportFormulas();




                case "RentRollActivityCompSummary":
                case "SubsidyRentRollReport":
                    return new SummaryColumnGenerator();




                case "ReportPayablesRegister":
                case "AgedPayables":
                case "AgedReceivables":
                    formulaGenerator = new FullTableFormulaGenerator();
                    formulaGenerator.SetDefenitionForBeyondFormulaRange(formulaGenerator.IsNonDataCell);
                    return formulaGenerator;




                case "CollectionsAnalysisSummary":
                    formulaGenerator = new FullTableFormulaGenerator();
                    formulaGenerator.SetDataCellDefenition(                                     //matches a percentage
                        cell => FormulaManager.IsDollarValue(cell) || Regex.IsMatch(cell.Text, "^\\d?\\d{2}([.]\\d{2})?%$"));


                    return formulaGenerator;



                case "RentRollPortfolio":
                    formulaGenerator = new FullTableFormulaGenerator();
                    double ignored;
                    formulaGenerator.SetDataCellDefenition(cell => FormulaManager.IsDollarValue(cell) || Double.TryParse(cell.Text, out ignored));
                    formulaGenerator.SetDefenitionForBeyondFormulaRange(formulaGenerator.IsNonDataCell);
                    return formulaGenerator;




                case "ReportAccountBalances":
                case "ReportTenantBal":
                case "ProfitAndLossStatementByPeriod":
                case "LedgerReport":
                case "RentRollAll": 
                case "RentRollActivity_New":
                case "RentRollActivity":
                case "TrialBalance":
                case "ReportCashReceiptsSummary":
                case "JournalLedger":
                case "AgedAccountsReceivable":
                case "CollectionsAnalysis":
                case "InvoiceRecurringReport":
                case "VendorInvoiceReport":
                case "UnitInvoiceReport":
                    return new FullTableFormulaGenerator();




                //These reports dont fit into any existing system
                //AgedAccountsReceivable (its original totals are incorrect)




                //Reports Im working on
                case "Budget":
                    return new MultiFormulaGenerator(new FullTableSummaryColumn(), new RowSegmentFormulaGenerator());
                case "RentRollCommercialItemized":



                //Reports I dont have
                case "RentRollActivityTotals":
                case "ReportEscalateCharges":
                




                //This report cannot get formulas because it does not include some necessary data
                case "PaymentsHistory":



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
            switch (reportName)
            {


                case "ProfitAndLossStatementByPeriod":
                    return new string[] { "Total Income", "Total Expense", "Net Operating Income~-Total Expense,Total Income",
                        "Net Income~Net Operating Income,-Total Expense" };




                case "BalanceSheetDrillthrough":
                    return new string[] { "Current Assets=Total Current Assets", "Fixed Asset=Total Fixed Asset", 
                        "Other Asset=Total Other Asset", "Current Liabilities=Total Current Liabilities",
                        "Liability=Total Liability", "Long Term Liability=Total Long Term Liability", 
                        "Equity=Total Equity", "Total Liabilities~Total Long Term Liability,Total Liability,Total Current Liabilities",
                        "Total Assets~Total Other Asset,Total Fixed Asset,Total Current Assets", 
                        "Total Liabilities And Equity~Total Equity,Total Liabilities" };




                case "ProfitAndLossComp":
                    return new string[] { "INCOME=Total Income", "EXPENSE=Total Expense", 
                        "Net Operating Income~Total Income,-Total Expense", "Net Income~Net Operating Income,-Total Expense" };



                case "RentRollActivity":
                case "RentRollActivity_New":
                    switch (worksheetNum)
                    {
                        case 0:
                            return new string[] { "Total:" };

                        case 1:
                            return new string[] { "Total For ([A-Z][a-z]+)( [A-Z][a-z]+)*:" };

                        default:
                            return new string[0];
                    }




                case "ReportCashReceiptsSummary":
                    switch (worksheetNum)
                    {
                        case 0:
                            return new string[] { "Total Tenant Receivables:", "Total Other Receivables:", 
                                "Total For (January|February|March|April|May|June|July|August|September|October|November|December) [12]\\d\\d\\d:~Total Tenant Receivables:,Total Other Receivables:",
                                "Total For Commons at White Marsh:~Total For (January|February|March|April|May|June|July|August|September|October|November|December) [12]\\d\\d\\d:" };

                        default:
                            return new string[0];
                    }




                case "ReportTenantBal":
                    return new string[] { "Total Open Charges:", 
                        "Balance:~Total Open Charges:,Total Future Charges:,Total Unallocated Payments:" };




                case "ProfitAndLossBudget":
                    return new string[] { "INCOME=Total Income", "EXPENSE=Total Expense", 
                        "Net Operating Income~Total Income,-Total Expense", 
                        "Net Income~-Total Expense,Total Income,Net Operating Income" };




                case "BalanceSheetComp":
                    return new string[] { "Current Assets=Total Current Assets", "Fixed Asset=Total Fixed Asset",
                        "Other Asset=Total Other Asset", "Current Liabilities=Total Current Liabilities",
                        "Liability=Total Liability", "Liabilities And Equity=Total Liabilities And Equity",
                        "Long Term Liability=Total Long Term Liability", "Equity=Total Equity", 
                        "Total Liabilities~Total Long Term Liability,Total Liability,Total Current Liabilities",
                        "Total Assets~Total Other Asset,Total Fixed Asset,Total Current Assets" };




                case "ChargesCreditsReport":
                    return new string[] { "Total: \\$(\\d\\d\\d,)*\\d?\\d?\\d[.]\\d\\d" };




                case "SubsidyRentRollReport":
                    return new string[] { 
                        "Current Tenant \\sPortion of the Rent,Current  Subsidy Portion of the Rent=>Current Monthly \\sContract Rent" };




                case "RentRollActivityCompSummary":
                    return new string[] { "-Opening A/R,Closing A/R=>A/R [+][(]-[)]" };




                case "RentRollHistory":
                    switch (worksheetNum)
                    {
                        case 1:
                            return new string[] { "Residential: \\$\\d+(,\\d\\d\\d)*[.]\\d\\d", "Total: \\$\\d+(,\\d\\d\\d)*[.]\\d\\d", };

                        default:
                            return new string[0];
                    }






                case "RentRollActivityItemized":
                case "RentRollActivityItemized_New":
                    return new string[] { "1r=(\\d{4})|([A-Z]\\d\\d)", "1Beg\\s+Balance", "1Charges", "1Adjustments",
                        "1Payments", "1End Balance", "1Change", "2Total:" };




                case "BalanceSheetPropBreakdown":
                    return new string[] { "Current Assets=Total Current Assets", "Fixed Asset=Total Fixed Asset",
                        "Other Asset=Total Other Asset", "Current Liabilities=Total Current Liabilities", 
                        "Long Term Liability=Total Long Term Liability", "Equity=Total Equity", 
                        "Total Assets~Total Other Asset,Total Fixed Asset,Total Current Assets", 
                        "Total Liabilities~Total Current Liabilities,Total Long Term Liability", 
                        "Total Liabilities And Equity~Total Equity,Total Liabilities" };




                case "VendorInvoiceReportWithJournalAccounts":
                    switch (worksheetNum)
                    {
                        case 5:
                            return new string[] { "Total:" };

                        default:
                            return new string[] { "Amount Owed", "Amount Paid", "Balance" };
                    }




                case "ReportCashReceipts":
                    return new string[] { "r=[A-Z]\\d{4}", "Charge Total", "Amount" };




                case "RentRollAllItemized":
                    switch (worksheetNum)
                    {

                        case 0:
                            return new string[] { "1r=[A-Z]-\\d\\d", "1Monthly Charge", "1Annual Charge", "2Total:" };

                        case 1:
                            return new string[] { "1r=[A-Z]-\\d\\d", "1Monthly Charge", "1Annual Charge", "2Total:", "3sheet0", "3sheet1" };

                        case 2:
                            return new string[] { "1Total:", "2Subtotals=Total:" };

                        default:
                            return new string[0];
                    }




                case "ProfitAndLossStatementDrillThrough":
                    return new string[] { "Expense=Total Expense", "Income=Total Income", 
                        "Net Operating Income~-Total Expense,Total Income", "Net Income~Net Operating Income,-Total Expense" };




                case "PayablesAccountReport":
                    return new string[] { "Pool Furniture=Total Pool Furniture", "Hallways=Total Hallways", 
                        "Garage=Total Garage", "Elevators=Total Elevators", "Clubhouse=Total Clubhouse", 
                        "Total Common Area CapEx~Total Pool Furniture,Total Hallways,Total Garage,Total Elevators,Total Clubhouse", "Total~Total Common Area CapEx", 
                        "Total:~Total Common Area CapEx" };




                case "ReportOutstandingBalance":
                    switch (worksheetNum)
                    {
                        case 0:
                            return new string[] { "1r=[A-Z0-9]+", "1Balance", "2Total For Commons at( [A-Z][a-z]+)+:" };

                        default:
                            return new string[] { "Total" };
                    }




                case "CollectionsAnalysis":
                case "ReportPayablesRegister":
                case "AgedAccountsReceivable":
                case "ReportAccountBalances":
                case "JournalLedger":
                case "CollectionsAnalysisSummary":
                case "AgedPayables":
                case "AgedReceivables":
                    return new string[] { "Total" };




                case "VacancyLoss":
                case "VendorInvoiceReport":
                case "InvoiceRecurringReport":
                case "UnitInvoiceReport":
                case "RentRollPortfolio":
                case "TrialBalance":
                case "RentRollAll":
                    return new string[] { "Total:" };




                case "ProfitAndLossStatementByJob":
                    return new string[] { "Income=Total Income", "Expense=Total Expense", 
                        "Net Income~Total Income,-Total Expense" };



                case "TrialBalanceVariance":
                    return new string[] { "Asset=Total Asset", "Liability=Total Liability", "Equity=Total Equity", 
                        "Income=Total Income", "Expense=Total Expense", "Total:~Total Expense,Total Income,Total Equity,Total Liability,Total Asset" };



                case "ProfitAndLossExtendedVariance":
                    return new string[] { "INCOME=Total Income", "EXPENSE=Total Expense",
                        "Net Operating Income~Total Income,-Total Expense", 
                        "Net Income~Net Operating Income,Total Income,-Total Expense" };



                case "LedgerReport":
                    return new string[] { "Total \\d+ - Prepaid Contracts" };



                //Reports with minor issues:
                //ProfitAndLossExtendedVariance
                //AgedAccountsReceivable
                //RentRollCommercialItemized
                //LedgerReport




                //these reports I'm still working on
                case "Budget":
                    return new string[] { "1Total", "2INCOME=TOTAL INCOME", "2EXPENSE=TOTAL EXPENSE" };
                    //FIXME: I am not sure what rows need formulas
                case "RentRollCommercialItemized": //not sure what Im supposed to be adding here


                // these reports I dont have
                case "RentRollActivityTotals":
                case "ReportEscalateCharges":




                // this report does not have the necessary columns/data to get a formula
                // for the time being this report gets no formulas
                case "PaymentsHistory": 



                default:
                    return new string[0];
            }
        }
    }
}
