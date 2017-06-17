using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Extensions;

namespace ExcelFormulaExpressionParser.Functions
{
    internal static class TextFunctions
    {
        public static Expression Left(Expression textExpression, Expression numCharsExpression = null)
        {
            int numChars = numCharsExpression?.LambdaInvoke<int>() ?? 1;
            string text = textExpression.LambdaInvoke<string>();

            return Expression.Constant(text.Substring(0, numChars));
        }
    }
}