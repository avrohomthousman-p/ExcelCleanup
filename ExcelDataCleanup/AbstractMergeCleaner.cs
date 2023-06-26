using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataCleanup
{

    /// <summary>
    /// Defines the basic layout for cleaning merges in an excel file
    /// </summary>
    internal abstract class AbstractMergeCleaner : IMergeCleaner
    {
        public void Unmerge(ExcelWorksheet worksheet)
        {
            FindTableBounds(worksheet);

            UnMergeMergedSections(worksheet);

            ResizeCells(worksheet);

            DeleteColumns(worksheet);
        }




        /// <summary>
        /// Finds the first row that is considered part of the table in the specified worksheet and saves the 
        /// row number to a local variable for later use
        /// </summary>
        /// <param name="worksheet">the worksheet we are working on</param>
        /// <exception cref="Exception">if the first row of the table couldnt be found</exception>
        protected abstract void FindTableBounds(ExcelWorksheet worksheet);



        /// <summary>
        /// Unmerges all the merged sections in the worksheet.
        /// </summary>
        /// <param name="worksheet">the worksheet we are currently cleaning</param>
        protected abstract void UnMergeMergedSections(ExcelWorksheet worksheet);



        /// <summary>
        /// Resizes all columns and rows that need a resize
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        protected abstract void ResizeCells(ExcelWorksheet worksheet);



        /// <summary>
        /// Deletes all empty/unwanted columns in the table
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        protected abstract void DeleteColumns(ExcelWorksheet worksheet);







        /* Some Abstract Utility Methods */


        /// <summary>
        /// Checks if the cell at the specified coordinates is a data cell
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the cell is a data cell and false otherwise</returns>
        protected abstract bool IsDataCell(ExcelRange cell);



        /// <summary>
        /// Checks if the specified cell is inside the table in the worksheet, and not a header 
        /// above the table
        /// </summary>
        /// <param name="cell">the cell whose location is being checked</param>
        /// <returns>true if the specified cell is inside a table and false otherwise</returns>
        protected abstract bool IsInsideTable(ExcelRange cell);



        /// <summary>
        /// Checks if the specified cell is a major header.
        /// 
        /// A major header is defined as a header that is above the table.
        /// </summary>
        /// <param name="cell">the cell being checked</param>
        /// <returns>true if the specified cell contains a major header, and false otherwise</returns>
        protected abstract bool IsMajorHeader(ExcelRange cell);



        /// <summary>
        /// Checks if the specified cell is considered a minor header.
        /// 
        /// A minor header is defined as a merge cell that contains non-data text and is inside the table.
        /// </summary>
        /// <param name="cells">the cells that we are checking</param>
        /// <returns>true if the specified cells are a minor header and false otherwise</returns>
        protected abstract bool IsMinorHeader(ExcelRange cells);





        /* Some Utility Methods With Implementations */


        /// <summary>
        /// Checks if a cell has no text
        /// </summary>
        /// <param name="currentCells">the cell that is being checked for text</param>
        /// <returns>true if there is no text in the cell, and false otherwise</returns>
        protected bool IsEmptyCell(ExcelRange currentCells)
        {
            return currentCells.Text == null || currentCells.Text.Length == 0;
        }



        /// <summary>
        /// Sets the PatternType, Color, Border, Font, and Horizontal Alingment of all the cells
        /// in the specifed range.
        /// </summary>
        /// <param name="currentCells">the cells whose style must be set</param>
        /// <param name="style">all the styles we should use</param>
        protected void SetCellStyles(ExcelRange currentCells, ExcelStyle style)
        {


            //Ensure the color of the cells dont get changed
            currentCells.Style.Fill.PatternType = style.Fill.PatternType;
            if (currentCells.Style.Fill.PatternType != ExcelFillStyle.None)
            {
                currentCells.Style.Fill.PatternColor.SetColor(
                    GetColorFromARgb(style.Fill.PatternColor.LookupColor()));

                currentCells.Style.Fill.BackgroundColor.SetColor(
                    GetColorFromARgb(style.Fill.BackgroundColor.LookupColor()));
            }


            currentCells.Style.Border = style.Border;
            currentCells.Style.Font.Bold = style.Font.Bold;
            currentCells.Style.Font.Size = style.Font.Size;
            currentCells.Style.Font.Name = style.Font.Name;
            currentCells.Style.Font.Scheme = style.Font.Scheme;
            currentCells.Style.Font.Charset = style.Font.Charset;


            //Not sure why but it only sucsessfully sets these settings if these 2 lines are NOT executed
            //currentCells.Style.WrapText = style.WrapText;
            //currentCells.Style.HorizontalAlignment = style.HorizontalAlignment;
        }



        /// <summary>
        /// Generates A Color Object from an ARGB string.
        /// </summary>
        /// <param name="argb">the argb code of the color needed</param>
        /// <returns>an instance of System.Drawing.Color that matches the specified argb code</returns>
        protected System.Drawing.Color GetColorFromARgb(String argb)
        {
            if (argb.StartsWith("#"))
            {
                argb = argb.Substring(1);
            }
            else if (argb.StartsWith("0x"))
            {
                argb = argb.Substring(2);
            }


            System.Drawing.Color color = System.Drawing.Color.FromArgb(
                int.Parse(argb, System.Globalization.NumberStyles.HexNumber));


            return color;
        }
    }
}
