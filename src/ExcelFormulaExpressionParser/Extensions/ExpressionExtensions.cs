using System;
using System.Linq.Expressions;

namespace ExcelFormulaExpressionParser.Extensions
{
    public static class ExpressionExtensions
    {
        public static T LambdaInvoke<T>(this Expression expression)
        {
            var value = Expression.Lambda(expression).Compile().DynamicInvoke();
            if (value == null)
            {
                return default(T);
            }

            if (value is T)
            {
                return (T)value;
            }

            return (T)Convert.ChangeType(value, typeof(T));
        }

        public static Expression LambdaInvoke(this Expression expression)
        {
            var value = Expression.Lambda(expression).Compile().DynamicInvoke();
            return Expression.Constant(value);
        }
    }
}