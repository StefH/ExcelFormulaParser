using ExcelFormulaExpressionParser.Models;
using System;
using System.Linq.Expressions;

namespace ExcelFormulaExpressionParser.Expressions
{
    internal class XRangeExpression : Expression
    {
        public readonly XRange XRange;

        private XRangeExpression(XRange range)
        {
            XRange = range;
        }

        public static XRangeExpression Create(XRange range)
        {
            return new XRangeExpression(range);
        }

        public override Type Type => typeof(XRange);
    }
}