using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompatableExcelCleaner
{

    /// <summary>
    /// An attempt to redo the ExcelIterator class in a way that will be more useful and less buggy
    /// 
    /// A collection of iterator functions to help iterate through an excel worksheet
    /// </summary>
    internal static class ExcelIterator2
    {
        public static readonly Tuple<int, int> SHIFT_UP = new Tuple<int, int>(-1, 0);
        public static readonly Tuple<int, int> SHIFT_DOWN = new Tuple<int, int>(1, 0);
        public static readonly Tuple<int, int> SHIFT_LEFT = new Tuple<int, int>(0, -1);
        public static readonly Tuple<int, int> SHIFT_RIGHT = new Tuple<int, int>(0, 1);


        public static ExcelWorksheet worksheet { get; set; }



        /// <summary>
        /// Iterates through every cell in the chosen direction until it reaches the end of the worksheet. After this
        /// operation the iterator will reference the last element in the worksheet (in the direction of the iteration).
        /// </summary>
        /// <param name="shift">the direction to move in, represented as a tuple with row change, and column change</param>
        /// <param name="row">Optional argument: the starting row</param>
        /// <param name="col">Optional argument: the starting column</param>
        /// <returns>the cells it iterates through</returns>
        public static IEnumerable<ExcelRange> GetCells(Tuple<int, int> shift, int row = 1, int col = 1)
        {
            EnsureWorksheetWasSet();

            while (!OutOfBounds(row, col))
            {
                ExcelRange cell = worksheet.Cells[row, col];

                yield return cell;

                row += shift.Item1;
                col += shift.Item2;
            }
        }




        /// <summary>
        /// Iterates through every cell in the chosen direction until it reaches the end of the worksheet or the 
        /// specified predicate evaluates to true. The cell that made the predicate true is NOT returned.
        /// </summary>
        /// <param name="shift">the direction to move in, represented as a tuple with row change, and column change</param>
        /// <param name="stopIf">the condition that signal we should stop iterating</param>
        /// <param name="row">Optional argument: the starting row</param>
        /// <param name="col">Optional argument: the starting column</param>
        /// <returns>the cells it iterates through</returns>
        public static IEnumerable<ExcelRange> GetCells(Tuple<int, int> shift, Predicate<ExcelRange> stopIf, int row = 1, int col = 1)
        {
            EnsureWorksheetWasSet();

            while (!OutOfBounds(row, col))
            {
                ExcelRange cell = worksheet.Cells[row, col];

                if (stopIf.Invoke(cell))
                {
                    break;
                }


                yield return cell;

                row += shift.Item1;
                col += shift.Item2;
            }
        }





        /// <summary>
        /// Iterates through every cell in the chosen direction until it reaches the end of the worksheet. The only difference
        /// between this method and the GetCells(shift) method is that this method returns the row and column of each cell, not
        /// the ExcelRange object.
        /// </summary>
        /// <param name="shift">the direction to move in, represented as a tuple with row change, and column change</param>
        /// <param name="row">Optional argument: the starting row</param>
        /// <param name="col">Optional argument: the starting column</param>
        /// <returns>the row and column of the cells it iterates through (stored in a tuple)</returns>
        public static IEnumerable<Tuple<int, int>> GetCellCoordinates(Tuple<int, int> shift, int row = 1, int col = 1)
        {
            foreach (ExcelRange cell in GetCells(shift, row, col))
            {
                yield return new Tuple<int, int>(cell.Start.Row, cell.Start.Column);
            }
        }




        /// <summary>
        /// Iterates through every cell in the chosen direction until it reaches the end of the worksheet of the 
        /// specified predicate evaluates to true. The only difference between this method and the 
        /// GetCells(shift, stopIf) method is that this method returns the row and column of each cell, not
        /// the ExcelRange object. Note: The cell that made the predicate true is NOT returned.
        /// </summary>
        /// <param name="shift">the direction to move in, represented as a tuple with row change, and column change</param>
        /// <param name="row">Optional argument: the starting row</param>
        /// <param name="col">Optional argument: the starting column</param>
        /// <returns>the row and column of the cells it iterates through (stored in a tuple)</returns>
        public static IEnumerable<Tuple<int, int>> GetCellCoordinates(Tuple<int, int> shift, Predicate<ExcelRange> stopIf, int row = 1, int col = 1)
        {
            foreach (ExcelRange cell in GetCells(shift, stopIf, row, col))
            {
                yield return new Tuple<int, int>(cell.Start.Row, cell.Start.Column);
            }
        }




        /// <summary>
        /// Iterates through the entire table, returing each cell as it goes.
        /// </summary>
        /// <param name="row">Optional Argument: the row where iteration should start</param>
        /// <param name="col">Optional Argument: the column where iteration should start</param>
        /// <returns>each cell in the table</returns>
        public static IEnumerable<ExcelRange> GetAllCellsInTable(int row = 1, int col = 1)
        {
            for (; row <= worksheet.Dimension.End.Row; row++)
            {
                for (; col <= worksheet.Dimension.End.Column; col++)
                {
                    yield return worksheet.Cells[row, col];
                }

                col = 1;
            }
        }




        /// <summary>
        /// Iterates through the entire table until it reaches the end or the specified predicate becomes true.
        /// Note: the cell that made the predicate true is not returned.
        /// </summary>
        /// <param name="stopIf">a function that tells the iterator when to stop iterating</param>
        /// <param name="row">Optional Argument: the row where iteration should start</param>
        /// <param name="col">Optional Argument: the column where iteration should start</param>
        /// <returns>each cell in the table</returns>
        public static IEnumerable<ExcelRange> GetAllCellsInTable(Predicate<ExcelRange> stopIf, int row = 1, int col = 1)
        {
            foreach (ExcelRange cell in GetAllCellsInTable(row, col))
            {
                if (stopIf.Invoke(cell))
                {
                    break;
                }

                yield return cell;
            }
        }




        /// <summary>
        /// Iterates through the entire table, returing each cell coordinates as it goes.
        /// </summary>
        /// <param name="row">Optional Argument: the row where iteration should start</param>
        /// <param name="col">Optional Argument: the column where iteration should start</param>
        /// <returns>each cell in the table as a tuple with row and column number</returns>
        public static IEnumerable<Tuple<int, int>> GetAllCoordinatesInTable(int row = 1, int col = 1)
        {
            for (; row <= worksheet.Dimension.End.Row; row++)
            {
                for (; col <= worksheet.Dimension.End.Column; col++)
                {
                    yield return new Tuple<int, int>(row, col);
                }

                col = 1;
            }
        }




        /// <summary>
        /// Iterates through the entire table until it reaches the end or the specified predicate becomes true.
        /// Note: the cell that made the predicate true is not returned.
        /// </summary>
        /// <param name="stopIf">a function that tells the iterator when to stop iterating</param>
        /// <param name="row">Optional Argument: the row where iteration should start</param>
        /// <param name="col">Optional Argument: the column where iteration should start</param>
        /// <returns>each cell in the table as a tuple with row and column number</returns>
        public static IEnumerable<Tuple<int, int>> GetAllCoordinatesInTable(Predicate<ExcelRange> stopIf, int row = 1, int col = 1)
        {
            foreach(ExcelRange cell in GetAllCellsInTable(row, col))
            {
                if (stopIf.Invoke(cell))
                {
                    break;
                }

                yield return new Tuple<int, int>(cell.Start.Row, cell.Start.Column);
            }
        }





        /// <summary>
        /// Iterates backwards through every cell in the table.
        /// </summary>
        /// <param name="row">Optional Argument: the row where iteration should start</param>
        /// <param name="col">Optional Argument: the column where iteration should start</param>
        /// <returns>each cell the iterator passed through</returns>
        public static IEnumerable<ExcelRange> GetAllCellsInTableReverse(int row = 1, int col = 1)
        {
            for (; row > 0; row--)
            {
                for (; col > 0; col--)
                {
                    yield return worksheet.Cells[row, col];
                }

                col = worksheet.Dimension.End.Column;
            }
        }





        /// <summary>
        /// Iterates and finds all cells in the table that match the specified predicate.
        /// </summary>
        /// <param name="isDesiredCell">a predicate that returns true if this is the cell that you are looking for</param>
        /// <param name="row">Optional Argument: the row where iteration should start</param>
        /// <param name="col">Optional Argument: the column where iteration should start</param>
        /// <returns>all cells found that match the predicate</returns>

        public static IEnumerable<ExcelRange> GetAllMatchingCells(Predicate<ExcelRange> isDesiredCell, int row = 1, int col = 1)
        {
            ExcelRange cell;


            for (; row <= worksheet.Dimension.End.Row; row++)
            {
                for (; col <= worksheet.Dimension.End.Column; col++)
                {
                    cell = worksheet.Cells[row, col];

                    if (isDesiredCell.Invoke(cell))
                    {
                        yield return cell;
                    }
                }


                col = 1;
            }
        }



        /// <summary>
        /// Iterates and finds all cells in the table that match the specified predicate.
        /// </summary>
        /// <param name="isDesiredCell">a predicate that returns true if this is the cell that you are looking for</param>
        /// <param name="row">Optional Argument: the row where iteration should start</param>
        /// <param name="col">Optional Argument: the column where iteration should start</param>
        /// <returns>all cells found that match the predicate in the form of (row, col) tuples</returns>

        public static IEnumerable<Tuple<int, int>> GetAllMatchingCoordinates(Predicate<ExcelRange> isDesiredCell, int row = 1, int col = 1)
        {
            foreach (ExcelRange cell in GetAllMatchingCells(isDesiredCell, row, col))
            {
                yield return new Tuple<int, int>(cell.Start.Row, cell.Start.Column);
            }
        }






        /// <summary>
        /// Checks if the worksheet was initialized
        /// </summary>
        /// <exception cref="NullReferenceException">if the worksheet was not initialized</exception>
        private static void EnsureWorksheetWasSet()
        {
            if(worksheet == null)
            {
                throw new NullReferenceException("Worksheet is null. Check that it in set before the iterator is used");
            }
        }



        /// <summary>
        /// Checks if the iterator is out of bounds of the worksheet
        /// </summary>
        /// <param name="row">the current row</param>
        /// <param name="col">the current column</param>
        /// <returns>true if the iterator has gone out of bounds, and false otherwise</returns>
        private static bool OutOfBounds(int row, int col)
        {
            return col < 1 || col > worksheet.Dimension.End.Column || row < 1 || row > worksheet.Dimension.End.Row;
        }
    }
}
