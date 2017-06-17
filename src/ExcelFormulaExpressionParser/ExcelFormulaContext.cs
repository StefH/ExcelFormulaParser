using ExcelFormulaExpressionParser.Models;
using ExcelFormulaParser;

namespace ExcelFormulaExpressionParser
{
    public class ExcelFormulaContext : IExcelFormulaContext
    {
        public XSheet Sheet { get; set; }
    }
}