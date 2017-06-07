using System;
using System.Linq;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Models;

namespace ExcelFormulaExpressionParser.Expressions
{
    internal static class LookupAndReferenceExpression
    {
        // =VLOOKUP(11.5, A2:B5001, 2)
        public static Expression VLookup(Expression search, XRange range, Expression column)
        {
            foreach (var keyCell in range.Cells.Where(c => c.Row == 1))
            {
                var keyExpression = range.Expressions[keyCell.Column - 1];
                if (keyExpression != null)
                {
                    var e = Expression.GreaterThanOrEqual(keyExpression, search);
                    bool match = (bool) Expression.Lambda(e).Compile().DynamicInvoke();

                    if (match)
                    {
                        int columnIndex = Convert.ToInt32(((ConstantExpression) column).Value);
                        var valueExpression = range.Expressions[keyCell.Column - 1 + columnIndex - 1];
                        return valueExpression;
                    }
                }
            }

            return null;
        }
    }
}