//using System;
//using System.Linq;
//using System.Linq.Expressions;
//using ExcelFormulaExpressionParser.Constants;
//using ExcelFormulaExpressionParser.Extensions;
//using ExcelFormulaExpressionParser.Models;
//using ExcelFormulaParser;
//using log4net;

//namespace ExcelFormulaExpressionParser.Functions
//{
//    internal static class LookupAndReferenceFunctions
//    {
//        private static readonly ILog Log = LogManager.GetLogger(typeof(LookupAndReferenceFunctions));

//        public static Expression VLookup(ExpressionParser parser, Expression search, XRange range, Expression column, Expression matchmodeExpression = null)
//        {
//            Log.InfoFormat("{0}", range.Address);
//            object matchmode = matchmodeExpression != null ? Expression.Lambda(matchmodeExpression).Compile().DynamicInvoke() : true;
//            bool approximateMatch;
//            if (matchmode is bool)
//            {
//                approximateMatch = (bool)matchmode;
//            }
//            else
//            {
//                approximateMatch = Convert.ToBoolean(matchmode);
//            }

//            foreach (var keyCell in range.Cells.Where(c => c.Column == range.Start.Column))
//            {
//                int row = keyCell.Row;
//                var keyExpression = keyCell.Expression ?? Parse(keyCell.ExcelFormula, parser.Level, range, parser.Finder);
//                if (keyExpression != null)
//                {
//                    Expression checkExpression = approximateMatch ? Expression.GreaterThanOrEqual(keyExpression, search) : Expression.Equal(keyExpression, search);
//                    bool match = checkExpression.LambdaInvoke<bool>();

//                    if (match)
//                    {
//                        int columnIndex = MathFunctions.ToInt(column).LambdaInvoke<int>();

//                        // return (previous) match
//                        int matchRow = !approximateMatch ? row : row - 1;
//                        var cell = range.Cells.FirstOrDefault(c => c.Row == matchRow && c.Column == range.Start.Column + columnIndex - 1);

//                        return cell != null ? (cell.Expression ?? Parse(cell.ExcelFormula, parser.Level, range, parser.Finder)) : Errors.NA;
//                    }
//                }
//            }

//            return Errors.NA;
//        }

//        private static Expression Parse(ExcelFormula formula, int level, XRange range, ICellFinder finder)
//        {
//            return new ExpressionParser(formula, level + 1, new ExcelFormulaContext { Sheet = range.Sheet }, finder).Parse();
//        }
//    }
//}