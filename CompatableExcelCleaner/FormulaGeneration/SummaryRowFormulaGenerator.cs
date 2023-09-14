﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;

namespace CompatableExcelCleaner
{
    /// <summary>
    /// Generates formulas that add up cells from anywhere else in the worksheet. Header data should be passed
    /// to this class in this format: "headerOfFormulaCell~header1,header2,header3" where headerOfFormula cell
    /// is the header before  (to the left of) the cell that needs the formula and the other comma seperated 
    /// headers are headers in front of (to the left of) cells that should be included in the sum. If needed,
    /// you can also specify that a given header be subtracted instead of added by putting a minus sign before
    /// the name of any of the headers.
    /// 
    /// Note: this class will NOT do all the formulas necessary on the worksheet, only the ones that cant be done by
    /// other systems becuase their cells are not near each other. This class should be used in addition to whatever other
    /// formula generator is appropriate for the report being cleaned.
    /// </summary>
    internal class SummaryRowFormulaGenerator : IFormulaGenerator
    {

        private IsDataCell dataCellDef = new IsDataCell(cell => FormulaManager.IsDollarValue(cell));


        /// <summary>
        /// Adds all formulas to the worksheet as specified by the metadata in the headers array
        /// </summary>
        /// <param name="worksheet">the worksheet in need of formulas</param>
        /// <param name="headers">headers to look for to tell us which cells to add up</param>
        public void InsertFormulas(ExcelWorksheet worksheet, string[] headers)
        {
            foreach(string header in headers)
            {
                //Ensure that the header was intended for this class and not the FormulaGenerator
                if (!FormulaManager.IsNonContiguousFormulaRange(header))
                {
                    continue;
                }


                int indexOfTilda = header.IndexOf('~');
                string formulaHeader = header.Substring(0, indexOfTilda);
                string[] dataCells = header.Substring(indexOfTilda + 1).Split(',');


                FillInFormulas(worksheet, formulaHeader, dataCells);
            }
        }



        /// <summary>
        /// Inserts a formula into the cell with the specified header.
        /// </summary>
        /// <param name="worksheet">the worksheet being given formulas</param>
        /// <param name="formulaHeader">the text that should be found near the cell requiring a formula</param>
        /// <param name="dataCells">headers pointing to cells that should be included in the formula</param>
        private void FillInFormulas(ExcelWorksheet worksheet, string formulaHeader, string[] dataCells)
        {
            ExcelIterator iter = new ExcelIterator(worksheet);

            ExcelRange formulaCell = iter.GetFirstMatchingCell(cell => FormulaManager.TextMatches(cell.Text, formulaHeader));

            if(formulaCell == null)
            {
                Console.WriteLine("Cell with text " + formulaHeader + " not found. Formula insertion failed.");
                return;
            }


            List< Tuple<int, bool>> dataRows = GetRowsToIncludeInFormula(worksheet, dataCells, formulaCell.Start.Row);


            var nextDataColumn = iter.GetCells(ExcelIterator.SHIFT_RIGHT);
            foreach (ExcelRange cell in nextDataColumn)
            {

                //if this isnt a data cell, skip it (dont put a formula here)
                if(!dataCellDef(cell))
                {
                    continue;
                }


                
                formulaCell = iter.GetCurrentCell();

                int dataColumn = iter.GetCurrentCol();


                //now add the formula to the cell
                formulaCell.Formula = BuildFormula(worksheet, dataRows, dataColumn);
                formulaCell.Style.Locked = true;

                Console.WriteLine("Cell " + formulaCell.Address + " has been given this formula: " + formulaCell.Formula);
            }
        }



        /// <summary>
        /// Gets all the row numbers of the cells that are to be included in the formula, and if they should be subtracted 
        /// instead of added.
        /// </summary>
        /// <param name="worksheet">the worksheet that is being given formulas</param>
        /// <param name="headers">the text that signals that this data cell should be part of the formula</param>
        /// <param name="rowOfFormula">the row number of the cell the formula will be placed in</param>
        /// <returns>
        /// a list of row numbers of the cells that should be part of the formula, and booleans that are true
        /// if that row should be subtracted instead of added
        /// </returns>
        private List<Tuple<int, bool>> GetRowsToIncludeInFormula(ExcelWorksheet worksheet, string[] headers, int rowOfFormula)
        {

            List<Tuple<string, bool>> headerAndIsSubtraction = ConvertArray(headers);

            List<Tuple<int, bool>> results = new List<Tuple<int, bool>>();


            //Create iterator that points to the location of the formula cell
            ExcelIterator iter = new ExcelIterator(worksheet, rowOfFormula, 1);

            foreach(ExcelRange cell in iter.FindAllCellsReverse())
            {
                //If we already found all the headers we need
                if(headerAndIsSubtraction.Count == 0)
                {
                    break;
                }

                //if the cell has a dollar value or is empty, it isnt a header, so we can skip it
                if(FormulaManager.IsEmptyCell(cell) || dataCellDef(cell))
                {
                    continue;
                }



                for(int i = 0; i < headerAndIsSubtraction.Count; i++)
                {
                    Tuple<string, bool> tup = headerAndIsSubtraction[i];

                    if(FormulaManager.TextMatches(cell.Text, tup.Item1))
                    {
                        results.Add(new Tuple<int, bool>(iter.GetCurrentRow(), tup.Item2));
                        headerAndIsSubtraction.RemoveAt(i);
                        break;
                    }
                }
            }


            return results;
        }




        /// <summary>
        /// Converts an array of headers into a list of Tuples storing headers without the leading minus and a bool that
        /// is true if that header used to have a minus sign.
        /// </summary>
        /// <param name="headers">the headers that are to be included in the formula being created</param>
        /// <returns>a list of each header and a bool isSubtraction (true if this row should be subtracted in the formula)</returns>
        private List<Tuple<string, bool>> ConvertArray(string[] headers)
        {
            return headers.Select(                                      
                    (text => {
                        if (text.StartsWith("-"))
                            return new Tuple<string, bool>(text.Substring(1), true);
                        else
                            return new Tuple<string, bool>(text, false);
                    }))
                .ToList();
        }




        /// <summary>
        /// Builds the actual formula that should be inserted into the worksheet
        /// </summary>
        /// <param name="worksheet">the worksheet getting formulas</param>
        /// <param name="rowData">an array of tuples with row numbers that should be included in the formula,
        /// and booleans stating if they should be subtracted</param>
        /// <param name="column">the column the formula is in</param>
        /// <returns>the formula that needs to be added to the cell as a string</returns>
        private string BuildFormula(ExcelWorksheet worksheet, List<Tuple<int, bool>> rowData, int column)
        {
            StringBuilder formula = new StringBuilder("SUM(");

            foreach (Tuple<int, bool> i in rowData)
            {

                if (i.Item2)
                {
                    formula.Append("-");
                }

                formula.Append(GetAddress(worksheet, i.Item1, column)).Append(",");

            }



            formula.Remove(formula.Length - 1, 1); //delete the trailing comma

            formula.Append(")");

            return formula.ToString();
        }



        public void SetDataCellDefenition(IsDataCell isDataCell)
        {
            this.dataCellDef = isDataCell;
        }




        /// <summary>
        /// Gets the address of a cell as it would be displayed in a formula
        /// </summary>
        /// <param name="worksheet">the worksheet the cell is in</param>
        /// <param name="row">the row the cell is in</param>
        /// <param name="col">the column the cell is in</param>
        /// <returns>the cell address</returns>
        private static string GetAddress(ExcelWorksheet worksheet, int row, int col)
        {
            return worksheet.Cells[row, col].Address;
        }
    }
}
