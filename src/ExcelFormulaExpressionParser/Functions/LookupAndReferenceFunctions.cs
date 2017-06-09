using System.Linq;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Constants;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaExpressionParser.Models;

namespace ExcelFormulaExpressionParser.Functions
{
    internal static class LookupAndReferenceFunctions
    {
        public static Expression VLookup(Expression search, XRange range, Expression column, Expression matchmode = null)
        {
            bool exactMatch = matchmode?.LambdaInvoke<bool>() ?? false;

            foreach (var keyCell in range.Cells.Where(c => c.Column == range.Start.Column))
            {
                int row = keyCell.Row;
                var keyExpression = keyCell.Expression;
                if (keyExpression != null)
                {
                    Expression checkExpression = exactMatch ? Expression.Equal(keyExpression, search) : Expression.GreaterThanOrEqual(keyExpression, search);
                    bool match = checkExpression.LambdaInvoke<bool>();

                    if (match)
                    {
                        int columnIndex = MathFunctions.ToInt(column).LambdaInvoke<int>();

                        return range.Cells.First(c => c.Row == row && c.Column == range.Start.Column + columnIndex - 1).Expression;
                    }
                }
            }

            return Errors.NA;
        }
    }
}