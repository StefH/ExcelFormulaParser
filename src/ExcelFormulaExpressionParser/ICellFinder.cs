using System.Collections.Generic;
using ExcelFormulaExpressionParser.Models;

namespace ExcelFormulaExpressionParser
{
    public interface ICellFinder
    {
        IEnumerable<XCell> Find(string sheetName, string address);
    }
}
