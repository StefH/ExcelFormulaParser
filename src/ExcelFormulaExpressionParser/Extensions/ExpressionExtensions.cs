using System;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Expressions;

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

        //public static Expression LambdaInvoke(this Expression expression)
        //{
        //    if (expression == null)
        //    {
        //        return null;
        //    }

        //    if (expression is ConstantExpression || expression is XArgExpression || expression is XRangeExpression)
        //    {
        //        return expression;
        //    }

        //    object value = Expression.Lambda(expression).Compile().DynamicInvoke();
        //    return Expression.Constant(value);
        //}

        public static void LambdaInvoke(ref Expression expression)
        {
            if (!(expression == null || expression is ConstantExpression || expression is XArgExpression || expression is XRangeExpression))
            {
                object value = Expression.Lambda(expression).Compile().DynamicInvoke();
                expression = Expression.Constant(value);
            }
        }
    }
}