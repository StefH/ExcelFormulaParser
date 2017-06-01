using System;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelFormulaParser.Expressions.Console
{
    public static class MathExpression
    {
        public static Expression Abs(Expression value)
        {
            return Expression.Call(null, FindMethod("Abs", new[] { typeof(double) }), value);
        }

        public static Expression Cos(Expression value)
        {
            return Expression.Call(null, FindMethod("Cos", new[] { typeof(double) }), value);
        }

        public static Expression Round(Expression number, Expression digits)
        {
            return Expression.Call(null, FindMethod("Round", new[] { typeof(double), typeof(int) }), number, Expression.Convert(digits, typeof(int)));
        }

        public static Expression Sin(Expression value)
        {
            return Expression.Call(null, FindMethod("Sin", new[] { typeof(double) }), value);
        }

        private static MethodInfo FindMethod(string name, Type[] types)
        {
            return typeof(Math).GetMethod(name, types);
        }
    }
}
