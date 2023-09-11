using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO;
using CompatableExcelCleaner;

namespace ExcelDataCleanup
{
    public class FileCleaner
    {




        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]




        public static void Main(string[] args)
        {
            string filepath = "";

            if (args != null && args.Count() > 0)
            {
                filepath = args[0];
            }

            else
            {

                // C:\Users\avroh\Downloads\ExcelProject\PayablesAccountReport_large.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\PayablesAccountReport_1Prop.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\ReportPayablesRegister.xlsx

                // C:\Users\avroh\Downloads\ExcelProject\ProfitAndLossStatementDrillthrough.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\AgedReceivables.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\LedgerExport.xlsx

                // C:\Users\avroh\Downloads\ExcelProject\TrialBalance.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\ProfitAndLossStatementByPeriod.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\testFile.xlsx

                // C:\Users\avroh\Downloads\ExcelProject\BalanceSheetComp_742023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\BalanceSheetDrillthrough_722023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\BankReconcilliation_722023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\PaymentsHistory_722023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\LedgerReport_722023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports\AgedReceivables_7102023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports\BalanceSheetComp_7102023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports\AdjustmentReportMult_7102023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-2\AdjustmentReport_7102023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-3\CashFlow_7182023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-3\ChargesCreditsReport_7182023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-3\ProfitAndLossBudget_7182023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-3\CreditCardStatement_7182023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-3\CollectionsAnalysisSummary_7182023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\AgedAccountsReceivable_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\BankReconcilliation_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\ChargesCreditsReport_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\BalanceSheetPropBreakdown_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\BalanceSheetComp_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\AgedReceivables_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-4\AgedPayables_7192023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\BalanceSheetComp_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\AgedAccountsReceivable_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\BalanceSheetDrillthrough_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\AdjustmentReport_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\AgedReceivables_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ProfitAndLossStatementDrillThrough_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\RentRollAll_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\InvoiceDetail_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\LedgerReport_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ReportTenantBal_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\RentRollActivity_New_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\RentRollActivityCompSummary_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\RentRollHistory_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ReportOutstandingBalance_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ReportCashReceiptsSummary_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ReportCashReceipts_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ProfitAndLossStatementByPeriod_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\PayablesAccountReport_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\PendingWebPayments_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ProfitAndLossBudget_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ProfitAndLossComp_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\ChargesCreditsReport_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\RentRollPortfolio_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-5\PreprintedLeasesReport_7232023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-6\ReportTenantSummary_7252023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-6\TenantDirectory_7252023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-6\VacancyLoss_7252023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-6\SubsidyRentRollReport_7252023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\system-reports-6\VendorInvoiceReportWithJournalAccounts_7252023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\JournalLedger_8222023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\RentRollActivity_New_8222023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\RentRollActivityItemized_New_8222023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\ReportAccountBalances_8222023.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\RentRollAll.xlsx
                // C:\Users\avroh\Downloads\ExcelProject\missing-reports\RentRollAllItemized_8222023.xlsx




                Console.WriteLine("Please enter the filepath of the Excel report you want to clean:");
                filepath = Console.ReadLine();

                /*
                OpenFileDialog dialog = new OpenFileDialog();
                if (DialogResult.OK == dialog.ShowDialog())
                {
                    filepath = dialog.FileName;
                }
                Console.WriteLine("Hello World!");
                */
            }


            string reportName = GetReportName(filepath);
            byte[] output = OpenXLSX(ConvertFileToBytes(filepath), reportName, true);
            SaveByteArrayAsFile(output, filepath.Replace(".xlsx", "_fixed.xlsx"));
            Console.WriteLine("Press Enter to exit");
            Console.Read();

        }



        /// <summary>
        /// Determans what kind of report we are cleaning based on the name of the report
        /// </summary>
        /// <param name="filename">the name of the report's file</param>
        /// <returns>the name of the report type, or an empty string if it could not be determaned</returns>
        private static string GetReportName(string filename)
        {

            int start = filename.LastIndexOf('\\') + 1;
            int length;



            Regex regex = new Regex("^.+(_\\d+)[.]xlsx$"); //matches if the report name ends with an underscore followed by numbers

            if (regex.IsMatch(filename))
            {
                length = filename.Length - start;
                length -= (filename.Length - filename.LastIndexOf('_')); //minus the number of characters after the file name
            }
            else
            {
                length = filename.Length - start - 5; //if we just need to remove the .xlsx at the end
            }




            return filename.Substring(start, length);
        }




        /// <summary>
        /// Opens the specified file and writes its contents to a byte array. This function is only needed for testing. In production
        /// the file itself will be passed in as a byte array, not as a filepath.
        /// </summary>
        /// <param name="filepath">the location of the file</param>
        /// <returns>a byte array with the contents of the file in it</returns>
        private static byte[] ConvertFileToBytes(string filepath)
        {
            FileInfo existingFile = new FileInfo(filepath);
            byte[] fileData = new byte[existingFile.Length];


            var fileStream = existingFile.Open(FileMode.Open);
            int bytesRead = 0;
            int bytesToRead = (int)existingFile.Length;
            while (bytesToRead > 0)
            {
                int justRead = fileStream.Read(fileData, bytesRead, bytesToRead);

                if (justRead == 0)
                {
                    break;
                }

                bytesRead += justRead;
                bytesToRead -= justRead;
            }


            fileStream.Close();
            return fileData;
        }




        /// <summary>
        /// Cleans an excel file
        /// </summary>
        /// <param name="sourceFile">the excel file in byte form</param>
        /// <param name="reportName">the file name of the original excel file</param>
        /// <param name="addFormulas">should be true if you also want formulas added to the report</param>
        /// <return>the excel file in byte form</return>
        public static byte[] OpenXLSX(byte[] sourceFile, string reportName, bool addFormulas=false)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            using (ExcelPackage package = new ExcelPackage(new MemoryStream(sourceFile)))
            {


                ExcelWorksheet worksheet;
                for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
                {
                    worksheet = package.Workbook.Worksheets[i];

                    //If the worksheet is empty, Dimension will be null
                    if(worksheet.Dimension == null)
                    {
                        package.Workbook.Worksheets.Delete(i);
                        i--;
                        continue;
                    }

                    CleanWorksheet(worksheet, reportName);
                }


                byte[] results;
                if (addFormulas)
                {
                    results = FormulaManager.AddFormulas(package.GetAsByteArray(), reportName);
                }
                else
                {
                    results = package.GetAsByteArray();
                }
                
                

                Console.WriteLine("Workbook Cleanup complete");


                return results;

            }
        }



        /// <summary>
        /// Does the standard cleanup on the specified worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet to be cleaned</param>
        /// <param name="reportName">the name of the report we are working on</param>
        public static void CleanWorksheet(ExcelWorksheet worksheet, string reportName)
        {

            DeleteHiddenRows(worksheet);


            RemoveAllHyperLinks(worksheet);


            RemoveAllMerges(worksheet, reportName);


            UnGroupAllRows(worksheet);


            CorrectCellDataTypes(worksheet);


            DoAdditionalCleanup(worksheet, reportName);

        }




        /// <summary>
        /// Deletes all hidden rows in the specified worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private static void DeleteHiddenRows(ExcelWorksheet worksheet)
        {

            var end = worksheet.Dimension.End;

            for (int row = end.Row; row >= 1; row--)
            {
                if (worksheet.Row(row).Hidden == true)
                {
                    worksheet.DeleteRow(row);
                    Console.WriteLine("Deleted Hidden Row : " + row);
                }
                else if(RowIsSafeToDelete(worksheet, row))
                {
                    worksheet.DeleteRow(row);
                    Console.WriteLine("Deleted Very Small Row : " + row);
                }
            }
        }



        /// <summary>
        /// Checks if a row is empty and really really small and therefore no data would be lost if it was deleted
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="rowNumber">the row being checked</param>
        /// <returns>true if the row is safe to delete becuase it has no data in it</returns>
        private static bool RowIsSafeToDelete(ExcelWorksheet worksheet, int rowNumber)
        {

            var row = worksheet.Row(rowNumber);
            if(row.Height >= 3)
            {
                return false;
            }



            //Check to see if the row is empty and can be deleted
            for (int colNumber = 1; colNumber <= worksheet.Dimension.Columns; colNumber++)
            {

                var cell = worksheet.Cells[rowNumber, colNumber];

                if(cell.Text != null && cell.Text.Length > 0) //if the cell has text in it (its not empty)
                {
                    return false; //unsafe to delete this row as it might have important text
                }
            }


            return true;

        }



        /// <summary>
        /// Removes all hyperlinks that are in any of the cells in the specified worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private static void RemoveAllHyperLinks(ExcelWorksheet worksheet)
        {
            var start = worksheet.Dimension.Start;
            var end = worksheet.Dimension.End;

            for (int row = end.Row; row >= start.Row; row--)
            {
                for (int col = start.Column; col <= end.Column; ++col)
                {

                    var cell = worksheet.Cells[row, col];
                    StripCellOfHyperLink(cell, row, col);

                }
            }
        }



        /// <summary>
        /// Removes hyperlinks in the specified Excel Cell if any are present.
        /// </summary>
        /// <param name="cell">the cell whose hyperlinks should be removed</param>
        /// <param name="row">the row the cell is in</param>
        /// <param name="col">the column the cell is in</param>
        private static void StripCellOfHyperLink(ExcelRange cell, int row, int col)
        {
            if (cell.Hyperlink != null)
            {
                Console.WriteLine("Row=" + row.ToString() + " Col=" + col.ToString() + " Hyperlink=" + cell.Hyperlink);
                var val = cell.Value;
                cell.Hyperlink = null;
                cell.Value = val;
            }
        }




        /// <summary>
        /// Manages the unmerging
        /// </summary>
        /// <param name="worksheet">the worksheet whose cells must be unmerged</param>
        /// <param name="reportName">the name of the type of report being cleaned</param>
        private static void RemoveAllMerges(ExcelWorksheet worksheet, string reportName)
        {

            IMergeCleaner mergeCleaner = ReportMetaData.ChoosesCleanupSystem(reportName, worksheet.Index);

            try
            {
                mergeCleaner.Unmerge(worksheet);
            }
            catch(InvalidDataException e)
            {
                Console.WriteLine("Warning: Report " + reportName + " cannot be processed by the primary merge cleaner.");
                Console.WriteLine("Consider adding it to the list of reports that use the backup system.");
                Console.WriteLine("Error Message: " + e.Message);

                mergeCleaner = new BackupMergeCleaner();
                mergeCleaner.Unmerge(worksheet);
            }
            
        }



        /// <summary>
        /// Ungroups all grouped columns so that excel should not display a colapse or expand
        /// button (plus button or minus button) on the left margin.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        private static void UnGroupAllRows(ExcelWorksheet worksheet)
        {


            for (int row = 1; row <= worksheet.Dimension.Rows; row++)
            {
                var currentRow = worksheet.Row(row);

                if (currentRow.OutlineLevel > 0)
                {
                    currentRow.OutlineLevel = 0;
                }
            }
        }




        /// <summary>
        /// Checks all cells in the worksheet for numbers that are being stored as text, and replaces them with the actual number.
        /// The purpose of this is to remove the excel warning that comes up when numbers are stored as text.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        private static void CorrectCellDataTypes(ExcelWorksheet worksheet)
        {
            for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                {

                    ExcelRange cell = worksheet.Cells[i, j];

                    

                    //Skip Empty Cells
                    if (cell.Text == null || cell.Text.Length == 0)
                    {
                        continue;
                    }

                    //Skip Cells that already contain numbers
                    if(cell.Value.GetType() != typeof(string))
                    {
                        continue;
                    }




                    double unused;

                    if (Double.TryParse(cell.Text, out unused))   //if it is not a dollar value, we want to keep it as a string
                    {

                        //Ignore the excel error that we have a number stored as a string
                        var error = worksheet.IgnoredErrors.Add(cell);
                        error.NumberStoredAsText = true;
                        continue;                                 //skip the formatting at the end of this if statement

                    }
                    else if (cell.Text.StartsWith("$") || (cell.Text.StartsWith("($") && cell.Text.EndsWith(")")))
                    {

                        bool isNegative = cell.Text.StartsWith("(");

                        cell.Style.Numberformat.Format = "$#,##0.00;($#,##0.00)";
                        cell.Value = Double.Parse(CleanDollarValue(cell.Text));

                        if (isNegative)
                        {
                            cell.Value = (double)cell.Value * -1;
                        }

                    }
                    else if (IsDateWith2DigitYear(cell.Text))
                    {
                        string fourDigitYear = cell.Text.Substring(0, 6) + "20" + cell.Text.Substring(6);
                        cell.SetCellValue(0, 0, fourDigitYear);
                        continue;
                    }
                    else if (IsPercentage(cell.Text))
                    {
                        cell.Style.Numberformat.Format = "#0\\.00%;(#0\\.00%)";

                        bool isNegative = cell.Text.StartsWith("(");

                        string cleanedText = CleanPercentage(cell.Text);
                        double percentAsNumber = Double.Parse(cleanedText);

                        if (isNegative)
                        {
                            percentAsNumber *= -1;
                        }

                        cell.SetCellValue(0, 0, percentAsNumber);
                    }
                    else
                    {
                        continue; //If this data cannot be coverted to a number, skip the formatting below
                    }


                    
                    //When the alignment is set to general, text is left aligned but numbers are right aligned.
                    //Therefore if we change from text to number and we want to maintain alignment, we need to 
                    //change to right aligned.
                    if (cell.Style.HorizontalAlignment.Equals(ExcelHorizontalAlignment.General))
                    {
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    }
                }
            }
        }




        /// <summary>
        /// Prepares text to be converted to a double by removing all commas, the preceding dollar sign, 
        /// and the surrounding parenthesis in the string if present.
        /// </summary>
        /// <param name="text">the text that should be cleaned</param>
        /// <returns>cleaned text that should be safe to parse to a double</returns>
        private static string CleanDollarValue(String text)
        {
            string replacementText = RemoveParenthesis(text);


            replacementText = replacementText.Substring(1);             //Remove $


            replacementText = replacementText.Replace(",", "");         //remove all commas


            return replacementText;
        }



        /// <summary>
        /// Removes any parenthesis surrounding the text
        /// </summary>
        /// <param name="text">the text needing cleaning</param>
        /// <returns>the same text without the parenthesis</returns>
        private static string RemoveParenthesis(string text)
        {
            if(text.StartsWith("(") && text.EndsWith(")"))
            {
                return text.Substring(1, text.Length - 2);
            }
            else
            {
                return text;
            }
        }




        /// <summary>
        /// Checks if the specified text stores a date with a 2 digit year
        /// </summary>
        /// <param name="text">the text in question</param>
        /// <returns>true if the text matches the pattern of a date with a 2 digit year, and false otherwise</returns>
        private static bool IsDateWith2DigitYear(string text)
        {
            Regex reg = new Regex("^\\d\\d/\\d\\d/\\d\\d$");
            return reg.IsMatch(text);
        }




        /// <summary>
        /// Checks if the specified text is a percentage
        /// </summary>
        /// <param name="text">the text being checked</param>
        /// <returns>true if the text is a percentage stored in text, or false otherwise</returns>
        private static bool IsPercentage(string text)
        {
            return Regex.IsMatch(text, "(100([.]00)?%)|([.]\\d\\d%)|(\\d{1,2}([.]\\d\\d)?%)");
        }



        /// <summary>
        /// Removes all non digit characters from text so it can be converted into a double.
        /// </summary>
        /// <param name="text">the text that needs to be cleaned</param>
        /// <returns>a string that can be safely converted to a double</returns>
        private static string CleanPercentage(string text)
        {
            string cleanedText = RemoveParenthesis(text);

            cleanedText = cleanedText.Substring(0, cleanedText.Length - 1); //remove % sign

            return cleanedText;
        }



        /// <summary>
        /// Exceutes all report specific cleanup that needs to be done
        /// </summary>
        /// <param name="worksheet">the worksheet that is being cleaned</param>
        /// <param name="reportName">the report that is being cleaned</param>
        private static void DoAdditionalCleanup(ExcelWorksheet worksheet, string reportName)
        {
            if(reportName == "RentRollHistory")
            {
                RepairRentRollHistory(worksheet);
            }


            if(ReportMetaData.NeedsSummaryCellsMoved(reportName, worksheet.Index))
            {
                MoveOutOfPlaceSummaryCells(worksheet);
            }
        }



        /// <summary>
        /// Moves all data cells in the last (rightmost) column over one cell to the left.
        /// This addresses a bug that leaves them one cell too far to the right.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        private static void MoveOutOfPlaceSummaryCells(ExcelWorksheet worksheet)
        {
            int col = worksheet.Dimension.End.Column;
            ExcelRange source, dest;
            
            for(int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                source = worksheet.Cells[row, col];
                dest = worksheet.Cells[row, col - 1];
                if (CellShouldBeMoved(source, dest))
                {
                    source.CopyStyles(dest);
                    dest.Value = source.Value;
                    source.Value = "";
                }
                
            }
        }



        /// <summary>
        /// Checks if it is safe to transfer the data from the specified source cell to the specified destination.
        /// </summary>
        /// <param name="source">the source of the data</param>
        /// <param name="destination">the cell the data will be moved to</param>
        /// <returns>true if it is safe to do the transfer or false otherwise</returns>
        private static bool CellShouldBeMoved(ExcelRange source, ExcelRange destination)
        {
            if(!source.Text.StartsWith("$") && !source.Text.StartsWith("($"))
            {
                return false;
            }

            Console.WriteLine("Other cell says " + destination.Text);
            return destination.Text == null || destination.Text.Length == 0;
        }



        
        /// <summary>
        /// Does all cleanup that is specific to the rent roll history report (which has some serious formatting problems)
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        private static void RepairRentRollHistory(ExcelWorksheet worksheet)
        {
            int sheetNum = worksheet.Index;

            if (sheetNum != 1)
            {
                return;
            }


            Tuple<int, int> rowRange = FindMoneySection(worksheet);
            int moneySectionTop = rowRange.Item1;
            int moneySectionBottom = rowRange.Item2;

            rowRange = FindOccupancySection(worksheet, moneySectionBottom + 1);
            int occupancySectionTop = rowRange.Item1;
            int occupancySectionBottom = rowRange.Item2;


            RemoveEmptySections(worksheet, moneySectionTop, moneySectionBottom);
            RemoveEmptySections(worksheet, occupancySectionTop, occupancySectionBottom);

            ResizeColumnsToDefault(worksheet);
        }




        /// <summary>
        /// Finds the start and end rows of the section of the worksheet that has financial data in it.
        /// This code is used to clean the RentRollHistory report ONLY.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <returns>a tuple with the start and end rows of the money section</returns>
        private static Tuple<int, int> FindMoneySection(ExcelWorksheet worksheet)
        {
            string MONTH_REGEX = "(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) (19|20)\\d\\d";
            Predicate<ExcelRange> isMonth = cell => FormulaManager.TextMatches(cell.Text, MONTH_REGEX);


            ExcelIterator iter = new ExcelIterator(worksheet);

            int start = iter.GetFirstMatchingCell(isMonth).Start.Row;

            int end = iter.GetCellCoordinates(ExcelIterator.SHIFT_DOWN, cell => FormulaManager.IsEmptyCell(cell)).Last().Item1;

            return new Tuple<int, int>(start, end);
        }




        /// <summary>
        /// Finds the start and end rows of the section of the worksheet that has data about occupancy in it.
        /// This code is used to clean the RentRollHistory report ONLY.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="startSearchAtRow">the row to start searching at, which should be just after the money section</param>
        /// <returns>a tuple with the start and end rows of the occupancy section</returns>
        private static Tuple<int, int> FindOccupancySection(ExcelWorksheet worksheet, int startSearchAtRow)
        {
            string MONTH_REGEX = "(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) (19|20)\\d\\d";
            Predicate<ExcelRange> isMonth = cell => FormulaManager.TextMatches(cell.Text, MONTH_REGEX);


            ExcelIterator iter = new ExcelIterator(worksheet, startSearchAtRow, 1);

            int start = iter.GetFirstMatchingCell(isMonth).Start.Row;

            int end = iter.GetCellCoordinates(ExcelIterator.SHIFT_DOWN, cell => FormulaManager.IsEmptyCell(cell)).Last().Item1;

            return new Tuple<int, int>(start, end);
        }



        /// <summary>
        /// The RentRollHistory report has some sections that are empty should be deleted.
        /// This code is used to clean the RentRollHistory report ONLY.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        private static void RemoveEmptySections(ExcelWorksheet worksheet, int startRow, int endRow)
        {
            ExcelRange allCells, sampleCell;
            int column = LastNonEmptyColumn(worksheet, startRow);
            while (column > 0)
            {
                sampleCell = worksheet.Cells[startRow + 1, column]; //for checking if there is text
                allCells = worksheet.Cells[startRow, column, endRow, column]; //the cells that must be deleted if empty

                if (FormulaManager.IsEmptyCell(sampleCell))
                {
                    allCells.Delete(eShiftTypeDelete.Left);
                    column++;
                }


                column--;
            }

        }




        /// <summary>
        /// Finds the last (leftmost) non-empty cell in the specified row, and returns the column its in.
        /// This code is used to clean the RentRollHistory report ONLY.
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="row">the row we should check on</param>
        /// <returns>the column number of the leftmost column containing a non-empty cell</returns>
        private static int LastNonEmptyColumn(ExcelWorksheet worksheet, int row)
        {
            ExcelIterator iter = new ExcelIterator(worksheet, row, worksheet.Dimension.End.Column);
            iter.SkipEmptyCells(ExcelIterator.SHIFT_LEFT);
            return iter.GetCurrentCol();
        }



        /// <summary>
        /// Resizes all columns in the worksheet to a defualt size, regaurdless of original size.
        /// This code is used to clean the RentRollHistory report ONLY.
        /// </summary>
        /// <param name="worksheet">the worksheet in need of resizing</param>
        private static void ResizeColumnsToDefault(ExcelWorksheet worksheet)
        {
            for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
            {
                worksheet.Column(i).Width = 11;
            }
        }



        /// <summary>
        /// Saves the specified byte array to a file.
        /// </summary>
        /// <param name="fileData">the byte array that should be saved to the file</param>
        /// <param name="filepath">the filepath of the file</param>
        private static void SaveByteArrayAsFile(byte[] fileData, string filepath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            using (ExcelPackage package = new ExcelPackage(new MemoryStream(fileData)))
            {
                package.SaveAs(filepath);
            }
        }
    }

}
