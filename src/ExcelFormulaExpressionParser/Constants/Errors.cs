using System.Linq.Expressions;

namespace ExcelFormulaExpressionParser.Constants
{
    internal static class Errors
    {
        public static readonly Expression Div0 = Expression.Constant("#DIV/0!");
        public static readonly Expression NA = Expression.Constant("#N/A");
        public static readonly Expression Name = Expression.Constant("#NAME?");
        public static readonly Expression Null = Expression.Constant("#NULL!");
        public static readonly Expression Num = Expression.Constant("#NUM!");
        public static readonly Expression Ref = Expression.Constant("#REF!");
        public static readonly Expression Value = Expression.Constant("#VALUE!");
    }
}