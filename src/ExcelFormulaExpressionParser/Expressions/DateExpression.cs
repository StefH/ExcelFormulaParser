using System;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Helpers;

namespace ExcelFormulaExpressionParser.Expressions
{
    internal static class DateExpression
    {
        public static Expression Month(Expression expression)
        {
            return Expression.Constant(Compile(expression).Month);
        }

        public static Expression Now()
        {
            return Expression.Constant(DateTimeHelpers.ToOADate(DateTime.UtcNow));
        }

        public static Expression Year(Expression expression)
        {
            return Expression.Constant(Compile(expression).Year);
        }

        private static DateTime Compile(Expression expression)
        {
            double value = (double) Expression.Lambda(expression).Compile().DynamicInvoke();

            return DateTimeHelpers.FromOADate(value);
        }
    }
}