using System;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelFormulaParser.Expressions.Console
{
    public static class MathExpression
    {
        public static Expression Round(Expression number, Expression digits)
        {
            var func = new Func<Expression, Expression, Expression>((n, d) => Expression.Call(null, FindMethod("Round", new[] { typeof(double), typeof(int) }), n, d));

            return func.Invoke(number, Expression.Convert(digits, typeof(int)));
        }

        public static Expression Sin(Expression value)
        {
            var func = new Func<Expression, Expression>(v => Expression.Call(null, FindMethod("Sin", new[] { typeof(double) }), v));

            return func.Invoke(value);
        }

        private static MethodInfo FindMethod(string name, Type[] types)
        {
            return typeof(Math).GetMethod(name, types);
        }
    }
}
