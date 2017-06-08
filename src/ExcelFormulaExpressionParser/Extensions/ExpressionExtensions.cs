using System.Linq.Expressions;

namespace ExcelFormulaExpressionParser.Extensions
{
    public static class ExpressionExtensions
    {
        public static T LambdaInvoke<T>(this Expression expression)
        {
            return (T)Expression.Lambda(expression).Compile().DynamicInvoke();
        }
    }
}
