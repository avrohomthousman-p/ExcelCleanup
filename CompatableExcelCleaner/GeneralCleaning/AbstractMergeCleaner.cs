﻿using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;


namespace ExcelDataCleanup
{

    /// <summary>
    /// Defines the basic layout for cleaning merges in an excel file
    /// </summary>
    internal abstract class AbstractMergeCleaner : IMergeCleaner
    {
        public virtual void Unmerge(ExcelWorksheet worksheet)
        {
            FindTableBounds(worksheet);

            UnMergeMergedSections(worksheet);

            ResizeCells(worksheet);

            DeleteColumns(worksheet);

            AdditionalCleanup(worksheet);
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



        /// <summary>
        /// Does all aditional cleanup that is needed for the specified report
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="reportName">the report being cleaned</param>
        protected abstract void AdditionalCleanup(ExcelWorksheet worksheet);




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
        protected virtual void SetCellStyles(ExcelRange currentCells, ExcelStyle style)
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
        protected virtual System.Drawing.Color GetColorFromARgb(String argb)
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



        /// <summary>
        /// Ensures that the data in the specified cell is stored as text, not as a number. This is usefull
        /// if you need to ensure that a date in the report header does not get displayed as hashtags if the 
        /// column is too small.
        /// </summary>
        /// <param name="cell">the cell whose data must be converted to text</param>
        protected virtual void ConvertContentsToText(ExcelRange cell)
        {
            cell.SetCellValue(0, 0, cell.Text);
        }



        /// <summary>
        /// Splits a header cell with more than one line of text, into multiple rows,
        /// one for each line of text. Note: this operation should be done AFTER unmerging the cell.
        /// </summary>
        /// <param name="worksheet">the worksheet currently being cleaned</param>
        /// <param name="cells">the header cell containing multi-line text</param>
        protected virtual void SplitHeaderIntoMultipleRows(ExcelWorksheet worksheet, ExcelRange cells)
        {
            if (!cells.Text.Contains("\n"))
            {
                return;
            }


            string[] linesOfText = cells.Text.Split('\n');

            int numNewRows = linesOfText.Length - 1;
            int startRow = cells.Start.Row;
            int endRow = startRow + numNewRows;

            worksheet.InsertRow(startRow + 1, numNewRows);


            for (int rowNum = startRow; rowNum <= endRow; rowNum++)
            {
                var currentCell = worksheet.Cells[rowNum, cells.Start.Column];
                int indexOfText = rowNum - startRow;
                currentCell.SetCellValue(0, 0, linesOfText[indexOfText]);

                cells.CopyStyles(currentCell);
            }
        }



        /// <summary>
        /// Ensures that any major header that ends up in column 2 or 3, gets moved to column 1 (if possible)
        /// </summary>
        /// <param name="worksheet">the worksheet that is being cleaned</param>
        /// <param name="firstDataRow">the row that marks the beginning of the data section. All cells above it are major headers</param>
        protected virtual void MoveMajorHeadersLeft(ExcelWorksheet worksheet, int firstDataRow)
        {
            int lastColumnBeingMoved = Math.Min(3, worksheet.Dimension.End.Column);

            for (int col = 2; col <= lastColumnBeingMoved; col++)
            {
                for(int row = 1; row < firstDataRow; row++)
                {
                    MoveHeaderIfNeeded(worksheet, row, col);
                }
            }
        }




        /// <summary>
        /// Moves the header found at the specified coordinates to the first column of the worksheet if possible 
        /// </summary>
        /// <param name="worksheet">the worksheet being cleaned</param>
        /// <param name="row">the row number of the header cell</param>
        /// <param name="col">the column number of the header cell</param>
        /// <returns>true if the header was moved sucsessfully, and false otherwise</returns>
        protected virtual bool MoveHeaderIfNeeded(ExcelWorksheet worksheet, int row, int col)
        {
            ExcelRange source = worksheet.Cells[row, col];
            ExcelRange dest = worksheet.Cells[row, 1];

            if (!IsEmptyCell(dest))
            {
                return false;
            }


            source.CopyStyles(dest);
            dest.Value = source.Value;
            source.Value = null;
            return true;
        }
    }
}

