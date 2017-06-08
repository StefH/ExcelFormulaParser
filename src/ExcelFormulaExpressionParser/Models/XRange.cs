using System.Linq.Expressions;

namespace ExcelFormulaExpressionParser.Models
{
    public class XRange
    {
        public XSheet Sheet { get; set; }

        public string Address { get; set; }

        public CellAddress Start { get; set; }

        public CellAddress End { get; set; }

        public XCell[] Cells { get; set; }
    }
}