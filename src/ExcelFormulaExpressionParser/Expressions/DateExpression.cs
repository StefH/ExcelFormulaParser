using System;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Compatibility;

namespace ExcelFormulaExpressionParser.Expressions
{
    internal static class DateExpression
    {
        public static Expression Abs(Expression value)
        {
            return Expression.Call(null, typeof(DateTime).FindMethod("Abs", new[] { typeof(double) }), value);
        }
    }
}