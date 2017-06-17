using System.Linq.Expressions;

namespace ExcelFormulaExpressionParser.Constants
{
    internal static class MathConstants
    {
        public static readonly Expression Constant0 = Expression.Constant(0d);
        public static readonly Expression Constant1 = Expression.Constant(1d);
        public static readonly Expression Constant10 = Expression.Constant(10d);
    }
}