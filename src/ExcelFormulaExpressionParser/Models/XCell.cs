using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Utils;
using ExcelFormulaParser;

namespace ExcelFormulaExpressionParser.Models
{
    public class XCell
    {
        public int Column { get; }

        public int Row { get; }

        public string Address { get; }

        public string FullAddress { get; set; }

        public object Value { get; set; }

        public string Formula { get; set; }

        public ExcelFormula ExcelFormula { get; set; }

        public Expression Expression { get; set; }

        public XRow XRow { get; }

        public XCell(XRow xRow, string address)
        {
            XRow = xRow;
            Address = address;

            Column = ExcelUtils.GetExcelColumnNumber(address);
            Row = xRow.Row;
        }
    }
}