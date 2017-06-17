using System;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaExpressionParser.Helpers;

namespace ExcelFormulaExpressionParser.Functions
{
    internal static class DateFunctions
    {
        public static Expression Date(Expression yearExpression, Expression monthExpression, Expression dayExpression)
        {
            int year = yearExpression.LambdaInvoke<int>();
            int month = monthExpression.LambdaInvoke<int>();
            int day = dayExpression.LambdaInvoke<int>();

            var date = new DateTime(year, month, day);
            return Expression.Constant(DateTimeHelpers.ToOADate(date));
        }

        public static Expression Day(Expression expression)
        {
            return Expression.Constant((double)Invoke(expression).Day);
        }

        public static Expression EDate(Expression dateExpression, Expression monthExpression)
        {
            DateTime date = Invoke(dateExpression);
            int month = monthExpression.LambdaInvoke<int>();

            var endDate = date.AddMonths(month);
            return Expression.Constant(DateTimeHelpers.ToOADate(endDate));
        }

        public static Expression EndOfMonth(Expression dateExpression, Expression monthExpression)
        {
            DateTime date = Invoke(dateExpression);
            int month = monthExpression.LambdaInvoke<int>();

            var endofMonthDate = new DateTime(date.Year, date.Month, 1).AddMonths(month + 1).AddDays(-1);
            return Expression.Constant(DateTimeHelpers.ToOADate(endofMonthDate));
        }

        public static Expression Month(Expression expression)
        {
            return Expression.Constant((double)Invoke(expression).Month);
        }

        public static Expression Now()
        {
            return Expression.Constant(DateTimeHelpers.ToOADate(DateTime.UtcNow));
        }

        public static Expression Today()
        {
            return Expression.Constant(DateTimeHelpers.ToOADate(DateTime.UtcNow.Date));
        }

        public static Expression Year(Expression expression)
        {
            return Expression.Constant((double)Invoke(expression).Year);
        }

        private static DateTime Invoke(Expression expression)
        {
            double value = expression.LambdaInvoke<double>();

            return DateTimeHelpers.FromOADate(value);
        }
    }
}