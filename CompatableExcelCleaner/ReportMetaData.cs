using ExcelDataCleanup;
using System;


namespace CompatableExcelCleaner
{

    /// <summary>
    /// Stores and makes accesible all meta data about reports, like what merge cleaner and formula generator to
    /// use.
    /// </summary>
    internal static class ReportMetaData
    {


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
    }
}
