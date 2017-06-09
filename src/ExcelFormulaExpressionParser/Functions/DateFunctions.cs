using System;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaExpressionParser.Helpers;

namespace ExcelFormulaExpressionParser.Functions
{
    internal static class DateFunctions
    {
        public static Expression Day(Expression expression)
        {
            return Expression.Constant((double)Invoke(expression).Day);
        }

        public static Expression Month(Expression expression)
        {
            return Expression.Constant((double) Invoke(expression).Month);
        }

        public static Expression Now()
        {
            return Expression.Constant(DateTimeHelpers.ToOADate(DateTime.UtcNow));
        }

        public static Expression Year(Expression expression)
        {
            return Expression.Constant((double) Invoke(expression).Year);
        }

        private static DateTime Invoke(Expression expression)
        {
            double value = expression.LambdaInvoke<double>();

            return DateTimeHelpers.FromOADate(value);
        }
    }
}