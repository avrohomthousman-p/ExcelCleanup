﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompatableExcelCleaner
{

    /// <summary>
    /// Helps iterate through an excel worksheet in any direction
    /// </summary>
    internal class ExcelIterator
    {
        public static readonly Tuple<int, int> SHIFT_UP = new Tuple<int, int>(-1, 0);
        public static readonly Tuple<int, int> SHIFT_DOWN = new Tuple<int, int>(1, 0);
        public static readonly Tuple<int, int> SHIFT_LEFT = new Tuple<int, int>(0, -1);
        public static readonly Tuple<int, int> SHIFT_RIGHT = new Tuple<int, int>(0, 1);



        private ExcelWorksheet worksheet;
        private int row;
        private int col;


        public ExcelIterator(ExcelWorksheet w)
        {
            this.worksheet = w;
            this.row = 1;
            this.col = 1;
        }



        public int GetCurrentRow()
        {
            return row;
        }



        public int GetCurrentCol()
        {
            return col;
        }



        public Tuple<int, int> GetCurrentLocation()
        {
            return new Tuple<int, int>(row, col);
        }



        /// <summary>
        /// Gets the cell the iterator is currently referencing. This will be the same as the cell most recently returned
        /// by any of the iteration methods (GetCells or GetCellCoordinates) unless SetCurrentLocation was called since
        /// using those methods.
        /// </summary>
        /// <returns>the cell the iterator is refrencing</returns>
        public ExcelRange GetCurrentCell()
        {
            return worksheet.Cells[row, col];
        }



        /// <summary>
        /// Sets the location the iterator is pointing to
        /// </summary>
        /// <param name="row">the row the iterator should point to</param>
        /// <param name="col">the column the iterator should point to</param>
        /// <exception cref="ArgumentOutOfRangeException">if the row or column are out of bounds</exception>
        public void SetCurrentLocation(int row, int col)
        {
            if(row < 1 || row > worksheet.Dimension.End.Row)
            {
                throw new ArgumentOutOfRangeException("Row " + row + " is out of range for this worksheet");
            }
            if(col < 1 || col > worksheet.Dimension.End.Column)
            {
                throw new ArgumentOutOfRangeException("Column " + col + " is out range for this worksheet");
            }



            this.row = row;
            this.col = col;
        }




        /// <summary>
        /// Iterates through every cell in the chosen direction until it reaches the end of the worksheet. After this
        /// operation the iterator will reference the last element in the worksheet (in the direction of the iteration).
        /// </summary>
        /// <param name="shift">the direction to move in, represented as a tuple with row change, and column change</param>
        /// <returns>the cells it iterates through</returns>
        public IEnumerable<ExcelRange> GetCells(Tuple<int, int> shift)
        {
            while (!OutOfBounds())
            {
                ExcelRange cell = worksheet.Cells[row, col];

                yield return cell;

                row += shift.Item1;
                col += shift.Item2;
            }


            //Undo final increment so the iterator should point to the last cell in the worksheet
            row -= shift.Item1;
            col -= shift.Item2;
        }



        /// <summary>
        /// Iterates through every cell in the chosen direction until it reaches the end of the worksheet of the 
        /// specified predicate evaluates to true. After this operation the iterator will reference the cell that
        /// made our predicate true or the last cell before the iterator would go out of bounds.
        /// </summary>
        /// <param name="shift">the direction to move in, represented as a tuple with row change, and column change</param>
        /// <param name="stopIf">the condition that signal we should stop iterating</param>
        /// <returns>the cells it iterates through</returns>
        public IEnumerable<ExcelRange> GetCells(Tuple<int, int> shift, Predicate<ExcelRange> stopIf)
        {
            while (!OutOfBounds())
            {
                ExcelRange cell = worksheet.Cells[row, col];

                if (stopIf.Invoke(cell))
                {
                    yield return cell;
                    break;
                }


                yield return cell;

                row += shift.Item1;
                col += shift.Item2;
            }


            //Undo final increment so that the iterator points to the last cell before we went out of bounds
            if (OutOfBounds())
            {
                row -= shift.Item1;
                col -= shift.Item2;
            }
        }





        /// <summary>
        /// Iterates through every cell in the chosen direction until it reaches the end of the worksheet. The only difference
        /// between this method and the GetCells(shift) method is that this method returns the row and column of each cell, not
        /// the ExcelRange object.
        /// </summary>
        /// <param name="shift">the direction to move in, represented as a tuple with row change, and column change</param>
        /// <returns>the row and column of the cells it iterates through (stored in a tuple)</returns>
        public IEnumerable<Tuple<int, int>> GetCellCoordinnates(Tuple<int, int> shift)
        {
            foreach(ExcelRange cell in GetCells(shift))
            {
                yield return new Tuple<int, int>(row, col);
            }
        }




        /// <summary>
        /// Iterates through every cell in the chosen direction until it reaches the end of the worksheet of the 
        /// specified predicate evaluates to true. The only difference between this method and the 
        /// GetCells(shift, stopIf) method is that this method returns the row and column of each cell, not
        /// the ExcelRange object.
        /// </summary>
        /// <param name="shift">the direction to move in, represented as a tuple with row change, and column change</param>
        /// <returns>the row and column of the cells it iterates through (stored in a tuple)</returns>
        public IEnumerable<Tuple<int, int>> GetCellCoordinates(Tuple<int, int> shift, Predicate<ExcelRange> stopIf)
        {
            foreach (ExcelRange cell in GetCells(shift, stopIf))
            {
                yield return new Tuple<int, int>(row, col);
            }
        }




        /// <summary>
        /// Continues iteration from current cell, and keeps moving untill the specified condition returns false.
        /// After this operation, the iterator will be referencing the element that made the condition false, or
        /// the last element in the worksheet if we went out of bounds.
        /// </summary>
        /// <param name="shift">the direction to iterate in, stored as a tuple of (row increment, col increment)</param>
        /// <param name="condition">the condition that tells the iterator to keep going</param>
        public void SkipWhile(Tuple<int, int> shift, Predicate<ExcelRange> condition)
        {

            bool keepGoing = true;
            ExcelRange cell;

            while(!OutOfBounds() && keepGoing)
            {
                cell = worksheet.Cells[row, col];
                keepGoing = condition.Invoke(cell);
                row += shift.Item1;
                col += shift.Item2;
            }



            //undo final increment so we reference the element that made the predicate false
            //or the last element in the worksheet
            row -= shift.Item1;
            col -= shift.Item2;
        }




        /// <summary>
        /// Convienece method to skip all empty cells and continue iteration after them. This method is
        /// equivilent to calling SkipWhile with a predicate that checks if the cell is empty.
        /// </summary>
        /// <param name="shift">the direction to iterate in</param>
        public void SkipEmptyCells(Tuple<int, int> shift)
        {
            SkipWhile(shift, cell => FormulaManager.IsEmptyCell(cell));
        }



        /// <summary>
        /// Checks if the iterator is out of bounds of the worksheet
        /// </summary>
        /// <returns>true if the iterator has gone out of bounds, and false otherwise</returns>
        private bool OutOfBounds()
        {
            return col > 0 && col <= worksheet.Dimension.End.Column && row > 0 && row <= worksheet.Dimension.End.Row;
        }
    }
}