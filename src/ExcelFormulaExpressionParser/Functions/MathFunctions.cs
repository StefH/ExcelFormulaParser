using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Compatibility;
using ExcelFormulaExpressionParser.Constants;
using ExcelFormulaExpressionParser.Models;

namespace ExcelFormulaExpressionParser.Functions
{
    internal static class MathFunctions
    {
        public static Expression Abs(Expression value)
        {
            return Expression.Call(null, typeof(Math).FindMethod("Abs", new[] { typeof(double) }), value);
        }

        public static Expression Cos(Expression value)
        {
            return Expression.Call(null, typeof(Math).FindMethod("Cos", new[] { typeof(double) }), value);
        }

        public static Expression Max(Expression value1, Expression value2)
        {
            return Expression.Call(null, typeof(Math).FindMethod("Max", new[] { typeof(double), typeof(double) }), value1, value2);
        }

        public static Expression Power(Expression number, Expression power)
        {
            return Expression.Power(number, power);
        }

        public static Expression Round(Expression number, Expression digits = null)
        {
            return Expression.Call(null, typeof(Math).FindMethod("Round", new[] { typeof(double), typeof(int) }), number, digits != null ? ToInt(digits) : Expression.Constant(0));
        }

        public static Expression Sin(Expression value)
        {
            return Expression.Call(null, typeof(Math).FindMethod("Sin", new[] { typeof(double) }), value);
        }

        public static Expression Sqrt(Expression value)
        {
            return Expression.Call(null, typeof(Math).FindMethod("Sqrt", new[] { typeof(double) }), value);
        }

        public static Expression Sum(IList<Expression> expressions, IList<XRange> ranges)
        {
            Expression left = MathConstants.Constant0;
            Expression right = MathConstants.Constant0;

            if (expressions.Any())
            {
                left = expressions.Aggregate(Expression.Add);
            }

            if (ranges.Any())
            {
                right = ranges.SelectMany(r => r.Cells).Select(c => c.Expression).Aggregate(Expression.Add);
            }

            return Expression.Add(left, right);
        }

        public static Expression Trunc(IList<Expression> expressions)
        {
            var truncateMethod = typeof(Math).FindMethod("Truncate", new[] { typeof(double) });

            Expression digits = expressions.Count == 1 ? MathConstants.Constant0 : expressions[1];
            Expression first = Expression.Multiply(expressions[0], Power(MathConstants.Constant10, digits));
            Expression truncate = Expression.Call(null, truncateMethod, first);

            return Expression.Divide(truncate, Power(MathConstants.Constant10, digits));
        }

        public static Expression ToInt(Expression value)
        {
            return Expression.Convert(value, typeof(int));
        }
    }
}
