using ExcelFormulaExpressionParser.Models;

namespace ExcelFormulaExpressionParser
{
    public interface ICellFinder
    {
        XRange Find(string sheetName, string address);
    }
}
