using ExcelFormulaExpressionParser.Models;

namespace ExcelFormulaExpressionParser
{
    public interface ICellFinder
    {
        XRange Find(XSheet sheet, string address);
    }
}
