using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelFormulaExpressionParser.Expressions
{
    internal static class LookupAndReferenceExpression
    {
        // =VLOOKUP(11.5, A2:B5001, 2)
        public static Expression VLookup(Expression search, IList<Expression> range, Expression column)
        {
            var value = Expression.Lambda(search).Compile().DynamicInvoke();

            // foreach ()

            return null;
        }
    }
}