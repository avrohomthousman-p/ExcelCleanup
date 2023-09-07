using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompatableExcelCleaner.FormulaGeneration
{
    /// <summary>
    /// A custom Formula that adds all cells in the range that do not have formulas
    /// </summary>
    public class AddAllNonFormulaCells : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            double total = 0;

            foreach(var arg in arguments)
            {
                if (arg.Value is ExcelRange cell)
                {
                    if (!FormulaManager.CellHasFormula(cell))
                    {
                        try
                        {
                            total += cell.GetValue<Double>();
                        }
                        catch(InvalidCastException e)
                        {
                            return new CompileResult(eErrorType.Value);
                        }
                        catch(FormatException e)
                        {
                            return new CompileResult(eErrorType.Value);
                        }
                    }
                }
                else
                {
                    return new CompileResult(eErrorType.Value);
                }
            }


            return new CompileResult(total, DataType.Decimal);
        }
    }
}
