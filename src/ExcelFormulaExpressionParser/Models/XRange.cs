using System.Linq.Expressions;

namespace ExcelFormulaExpressionParser.Models
{
    internal class XRange
    {
        public string Sheet { get; set; }

        public string Range { get; set; }

        public Expression[] Expressions { get; set; }
    }
}