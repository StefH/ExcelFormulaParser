using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace ExcelFormulaExpressionParser.Expressions
{
    internal static class LogicalExpression
    {
        public static Expression And(IList<Expression> expressions)
        {
            return expressions.Aggregate(Expression.And);
        }

        public static Expression Or(IList<Expression> expressions)
        {
            return expressions.Aggregate(Expression.Or);
        }
    }
}