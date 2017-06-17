using System;
using System.Linq;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Constants;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaExpressionParser.Models;

namespace ExcelFormulaExpressionParser.Functions
{
    internal static class LookupAndReferenceFunctions
    {
        public static Expression VLookup(Expression search, XRange range, Expression column, Expression matchmodeExpression = null)
        {
            object matchmode = matchmodeExpression != null ? Expression.Lambda(matchmodeExpression).Compile().DynamicInvoke() : true;
            bool approximateMatch;
            if (matchmode is bool)
            {
                approximateMatch = (bool)matchmode;
            }
            else
            {
                approximateMatch = Convert.ToBoolean((double)matchmode);
            }

            foreach (var keyCell in range.Cells.Where(c => c.Column == range.Start.Column))
            {
                int row = keyCell.Row;
                var keyExpression = keyCell.Expression;
                if (keyExpression != null)
                {
                    Expression checkExpression = approximateMatch ? Expression.GreaterThanOrEqual(keyExpression, search) : Expression.Equal(keyExpression, search);
                    bool match = checkExpression.LambdaInvoke<bool>();

                    if (match)
                    {
                        int columnIndex = MathFunctions.ToInt(column).LambdaInvoke<int>();

                        // return (previous) match
                        int matchRow = !approximateMatch ? row : row - 1;
                        var cell = range.Cells.FirstOrDefault(c => c.Row == matchRow && c.Column == range.Start.Column + columnIndex - 1);

                        return cell != null ? cell.Expression : Errors.NA;
                    }
                }
            }

            return Errors.NA;
        }
    }
}