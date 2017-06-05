using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace ExcelFormulaExpressionParser.Expressions
{
    internal static class LogicalExpression
    {
        public static Expression And(IList<Expression> expressions)
        {
            Expression result = expressions[0];
            foreach (var expression in expressions.Skip(1))
            {
                result = Expression.And(result, expression);
            }

            return result;
        }

        public static Expression Or(IList<Expression> expressions)
        {
            Expression result = expressions[0];
            foreach (var expression in expressions.Skip(1))
            {
                result = Expression.Or(result, expression);
            }

            return result;
        }
    }
}