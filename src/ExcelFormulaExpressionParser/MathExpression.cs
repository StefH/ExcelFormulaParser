using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Compatibility;

namespace ExcelFormulaExpressionParser
{
    public static class MathExpression
    {
        private static readonly Expression Constant0 = Expression.Constant(0d);
        private static readonly Expression Constant10 = Expression.Constant(10d);

        public static Expression Abs(Expression value)
        {
            return Expression.Call(null, typeof(Math).FindMethod("Abs", new[] { typeof(double) }), value);
        }

        public static Expression Cos(Expression value)
        {
            return Expression.Call(null, typeof(Math).FindMethod("Cos", new[] { typeof(double) }), value);
        }

        public static Expression Power(Expression number, Expression power)
        {
            return Expression.Power(number, power);
        }

        public static Expression Round(Expression number, Expression digits)
        {
            return Expression.Call(null, typeof(Math).FindMethod("Round", new[] { typeof(double), typeof(int) }), number, ToInt(digits));
        }

        public static Expression Sin(Expression value)
        {
            return Expression.Call(null, typeof(Math).FindMethod("Sin", new[] { typeof(double) }), value);
        }

        public static Expression Sqrt(Expression value)
        {
            return Expression.Call(null, typeof(Math).FindMethod("Sqrt", new[] { typeof(double) }), value);
        }

        public static Expression Sum(IList<Expression> expressions)
        {
            Expression result = expressions[0];
            foreach (var expression in expressions.Skip(1))
            {
                result = Expression.Add(result, expression);
            }

            return result;
        }

        public static Expression Trunc(IList<Expression> expressions)
        {
            var truncateMethod = typeof(Math).FindMethod("Truncate", new[] {typeof(double)});

            Expression digits = expressions.Count == 1 ? Constant0 : expressions[1];
            Expression first = Expression.Multiply(expressions[0], Power(Constant10, digits));
            Expression truncate = Expression.Call(null, truncateMethod, first);

            return Expression.Divide(truncate, Power(Constant10, digits));
        }

        private static Expression ToInt(Expression value)
        {
            return Expression.Convert(value, typeof(int));
        }
    }
}
