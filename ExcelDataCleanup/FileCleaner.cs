using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


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

            

            OpenXLSX( ConvertFileToBytes(filepath), filepath );
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
            int bytesToRead = (int) existingFile.Length;
            while (bytesToRead > 0)
            {
                int justRead = fileStream.Read(fileData, bytesRead, bytesToRead);

                if(justRead == 0)
                {
                    break;
                }

                bytesRead += justRead;
                bytesToRead -= justRead;
            }


            return fileData;
        }




        /// <summary>
        /// Opens an existing excel file and reads some values and properties
        /// </summary>
        /// <param name="file">the excel file in byte form</param>
        /// <param name="originalFileName">the file name of the original excel file</param>
        public static void OpenXLSX(byte[] file, string originalFileName)
        {
            string reportName = GetReportName(originalFileName);


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            using (ExcelPackage package = new ExcelPackage(new MemoryStream(file)))
            {
                //Get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];


                DeleteHiddenRows(worksheet);


                RemoveAllHyperLinks(worksheet);


                RemoveAllMerges(worksheet, reportName);


                FixExcelTypeWarnings(worksheet);





                package.SaveAs(originalFileName.Replace(".xlsx", "_fixed.xlsx"));

            }

            Console.WriteLine("Workbook Cleanup complete");
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

            int length = filename.Length - start - 5; //we dont want the .xlsx at the end

            return filename.Substring(start, length);
        }




        /// <summary>
        /// Deletes all hidden rows in the specified worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        private static void DeleteHiddenRows(ExcelWorksheet worksheet)
        {
            var start = worksheet.Dimension.Start;
            var end = worksheet.Dimension.End;

            for (int row = end.Row; row >= start.Row; row--)
            {
                if (worksheet.Row(row).Hidden == true)
                {
                    worksheet.DeleteRow(row);
                    Console.WriteLine("Deleted Hidden Row : " + row);
                }
            }
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
                //worksheet.Cells[cell.EntireColumn.ToString()].Merge = false;
                //cell.Hyperlink.ReferenceAddress("");

                Console.WriteLine("Row=" + row.ToString() + " Col=" + col.ToString() + " Hyperlink=" + cell.Hyperlink);
                //  Uri uval = new Uri(cell.Text, UriKind.Relative);
                // cell.Hyperlink;
                var val = cell.Value;
                cell.Hyperlink = null;
                ////cell.Hyperlink = new Uri(cell.ToString(), UriKind.Absolute);
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

            IMergeCleaner mergeCleaner = ChoosesCleanupSystem(reportName);
            //IMergeCleaner mergeCleaner = new PrimaryMergeCleaner();
            //IMergeCleaner mergeCleaner = new BackupMergeCleaner();

            mergeCleaner.Unmerge(worksheet);
        }



        /// <summary>
        /// Chosses the version of merge cleanup code that would work best for the specified report
        /// </summary>
        /// <param name="reportType">the type of report that needs unmerging</param>
        /// <returns>an instance of IMergeCleaner that should be used to clean the report</returns>
        private static IMergeCleaner ChoosesCleanupSystem(string reportType)
        {
            switch (reportType)
            {
                case "TrialBalance":
                case "ProfitAndLossStatementDrillthrough":
                    return new BackupMergeCleaner();

                default:
                    return new PrimaryMergeCleaner();
            }
        }




        /// <summary>
        /// Checks all cells in the worksheet for numbers that are being stored as text, and replaces them with the actual number.
        /// The purpose of this is to remove the excel warning that comes up when numbers are stored as text.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        private static void FixExcelTypeWarnings(ExcelWorksheet worksheet)
        {
            for (int i = 1; i <= worksheet.Dimension.Rows; i++)
            {
                for (int j = 1; j <= worksheet.Dimension.Columns; j++)
                {

                    ExcelRange cell = worksheet.Cells[i, j];

                    double? data = ConvertToNumber(cell.Text);

                    if (data != null)
                    {

                        cell.Value = data; //Replace the cell data with the same thing just not in text form


                        //When the alingment is set to general, text is left aligned but numbers are right aligned.
                        //Therefore if we change from text to number and we want to maintain alignment, we need to 
                        //change to right aligned.
                        if (cell.Style.HorizontalAlignment.Equals(ExcelHorizontalAlignment.General))
                        {
                            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }
                    }
                }
            }
        }



        /// <summary>
        /// Attempts to convert that specified string into a double
        /// </summary>
        /// <param name="data">the text that should be converted to a number</param>
        /// <returns>the text as a double object or null if it could not be converted</returns>
        private static double? ConvertToNumber(string data)
        {

            double result;

            bool sucsess = Double.TryParse(data, out result);


            return (sucsess ? result : null);
        }

    }

}
