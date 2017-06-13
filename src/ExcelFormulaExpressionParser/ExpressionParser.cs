using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Functions;
using ExcelFormulaExpressionParser.Models;
using ExcelFormulaExpressionParser.Utils;
using ExcelFormulaParser;
using JetBrains.Annotations;
using TokenType = ExcelFormulaParser.ExcelFormulaTokenType;
using TokenSubtype = ExcelFormulaParser.ExcelFormulaTokenSubtype;
using ExcelFormulaExpressionParser.Expressions;

namespace ExcelFormulaExpressionParser
{
    public class ExpressionParser
    {
        private int _index;

        private ExcelFormulaToken CT => _list[_index];

        private readonly IList<ExcelFormulaToken> _list;
        private readonly ExcelFormulaContext _context;
        private readonly ICellFinder _finder;

        /// <summary>
        /// ExpressionParser
        /// </summary>
        /// <param name="tokens">The ExcelFormula or a list from ExcelFormulaTokens.</param>
        public ExpressionParser([NotNull] IList<ExcelFormulaToken> tokens) :
            this(tokens, null, (ICellFinder)null)
        {
        }

        /// <summary>
        /// ExpressionParser
        /// </summary>
        /// <param name="tokens">The ExcelFormula or a list from ExcelFormulaTokens.</param>
        /// <param name="context">The ExcelFormulaContext. (Optional if no real Excel Workbook is parsed.)</param>
        /// <param name="sheets">The XSheets from a Excel Workbook. (Optional if no real Excel Workbook is parsed.)</param>
        public ExpressionParser([NotNull] IList<ExcelFormulaToken> tokens, [CanBeNull] ExcelFormulaContext context = null, [CanBeNull] List<XSheet> sheets = null) :
            this(tokens, context, sheets != null ? new CellFinder(sheets) : null)
        {
        }

        /// <summary>
        /// ExpressionParser
        /// </summary>
        /// <param name="tokens">The ExcelFormula or a list from ExcelFormulaTokens.</param>
        /// <param name="context">The ExcelFormulaContext. (Optional if no real Excel Workbook is parsed.)</param>
        /// <param name="finder">The cellfinder to finds cells in a workbook. (Optional if no real Excel Workbook is parsed.)</param>
        public ExpressionParser([NotNull] IList<ExcelFormulaToken> tokens, [CanBeNull] ExcelFormulaContext context = null, [CanBeNull] ICellFinder finder = null)
        {
            _list = tokens;
            _context = context;
            _finder = finder;
        }

        public Expression Parse()
        {
            _index = 0;

            return ParseArgs();
        }

        private void Next()
        {
            if (_index + 1 < _list.Count)
            {
                _index++;
            }
        }

        private Expression ParseArgs()
        {
            Expression left = ParseAdditive();

            while (CT.Type == TokenType.Argument)
            {
                Next();
                Expression right = ParseAdditive();

                var argExpression = left as XArgExpression;
                if (argExpression != null)
                {
                    left = argExpression.Add(right);
                }
                else
                {
                    argExpression = XArgExpression.Create(left);
                    left = argExpression.Add(right);
                }
            }

            return left;
        }

        // +, -
        private Expression ParseAdditive()
        {
            Expression left = ParseMultiplication();

            while (CT.Type == TokenType.OperatorInfix && CT.Subtype == TokenSubtype.Math && (CT.Value == "+" || CT.Value == "-"))
            {
                var op = CT;
                Next();
                Expression right = ParseMultiplication();

                switch (op.Value)
                {
                    case "+":
                        left = Expression.Add(left, right);
                        break;
                    case "-":
                        left = Expression.Subtract(left, right);
                        break;
                }
            }

            return left;
        }

        // *, /, ^
        private Expression ParseMultiplication()
        {
            Expression left = ParseLogical();

            while (CT.Type == TokenType.OperatorInfix && CT.Subtype == TokenSubtype.Math && (CT.Value == "^" || CT.Value == "*" || CT.Value == "/"))
            {
                var op = CT;
                Next();
                Expression right = ParseLogical();

                switch (op.Value)
                {
                    case "^":
                        left = MathFunctions.Power(left, right);
                        break;
                    case "*":
                        left = Expression.Multiply(left, right);
                        break;
                    case "/":
                        left = Expression.Divide(left, right);
                        break;
                    default:
                        throw new NotSupportedException(op.Value);
                }
            }

            return left;
        }

        // >, >=, <, <=, <>, =
        private Expression ParseLogical()
        {
            Expression left = ParseFunction();

            while (CT.Type == TokenType.OperatorInfix && CT.Subtype == TokenSubtype.Logical)
            {
                var op = CT;
                Next();
                Expression right = ParseFunction();

                switch (op.Value)
                {
                    case ">":
                        left = Expression.GreaterThan(left, right);
                        break;
                    case ">=":
                        left = Expression.GreaterThanOrEqual(left, right);
                        break;
                    case "<":
                        left = Expression.LessThan(left, right);
                        break;
                    case "<=":
                        left = Expression.LessThanOrEqual(left, right);
                        break;
                    case "<>":
                        left = Expression.NotEqual(left, right);
                        break;
                    case "=":
                        left = Expression.Equal(left, right);
                        break;
                    default:
                        throw new NotSupportedException(op.Value);
                }
            }

            return left;
        }

        // =ROUND(-3 * ROUND(1.001 * 2, 1) + 7.001, 1) + 11
        private Expression ParseFunction()
        {
            Expression left = ParseRange();

            while (CT.Type == TokenType.Function || CT.Type == TokenType.Subexpression)
            {
                var op = CT;
                Next();

                int indent = 0;
                var tokens = new List<ExcelFormulaToken>();
                while (!(CT.Subtype == TokenSubtype.Stop && indent == 0))
                {
                    if (CT.Subtype == TokenSubtype.Start)
                    {
                        indent++;
                    }

                    if (CT.Subtype == TokenSubtype.Stop)
                    {
                        indent--;
                    }

                    tokens.Add(CT);

                    Next();
                }

                Next();
                var right = ParseRange();

                // If there is more to be done (right != null), return right, else return the value from this function.
                return right ?? ParseFunctionWithArgs(op.Value, tokens, _context);
            }

            return left;
        }

        private Expression ParseFunctionWithArgs(string functionName, IList<ExcelFormulaToken> tokens, ExcelFormulaContext context)
        {
            void AddToArgumentsList(ICollection<object> args, Expression expression)
            {
                var argExpression = expression as XArgExpression;
                if (argExpression != null)
                {
                    foreach (var xargExpression in argExpression.Expressions)
                    {
                        AddToArgumentsList(args, xargExpression);
                    }

                    return;
                }

                var rangeExpression = expression as XRangeExpression;
                if (rangeExpression != null)
                {
                    args.Add(rangeExpression.XRange);
                    return;
                }

                args.Add(expression);
            }

            var arguments = new List<object>();

            if (tokens.Any())
            {
                AddToArgumentsList(arguments, Parse(tokens, context));
            }

            var expressions = arguments.Where(a => a is Expression).Cast<Expression>().ToArray();
            var xranges = arguments.Where(a => a is XRange).Cast<XRange>().ToArray();

            switch (functionName)
            {
                case "": return expressions[0];
                case "ABS": return MathFunctions.Abs(expressions[0]);
                case "AND": return LogicalFunctions.And(expressions);
                case "COS": return MathFunctions.Cos(expressions[0]);
                case "DATE": return DateFunctions.Date(expressions[0], expressions[1], expressions[2]);
                case "DAY": return DateFunctions.Day(expressions[0]);
                case "IF": return Expression.Condition(expressions[0], expressions[1], expressions[2]);
                case "MONTH": return DateFunctions.Month(expressions[0]);
                case "NOW": return DateFunctions.Now();
                case "OR": return LogicalFunctions.Or(expressions);
                case "POWER": return MathFunctions.Power(expressions[0], expressions[1]);
                case "ROUND": return MathFunctions.Round(expressions[0], expressions.Length == 2 ? expressions[1] : null);
                case "SIN": return MathFunctions.Sin(expressions[0]);
                case "SQRT": return MathFunctions.Sqrt(expressions[0]);
                case "SUM": return MathFunctions.Sum(expressions, xranges);
                case "TRUNC": return MathFunctions.Trunc(expressions);
                case "VLOOKUP": return LookupAndReferenceFunctions.VLookup(expressions[0], xranges[0], expressions[1], expressions.Length == 3 ? expressions[2] : null);
                case "YEAR": return DateFunctions.Year(expressions[0]);

                default:
                    throw new NotImplementedException(functionName);
            }
        }

        private Expression ParseRange()
        {
            Expression left = ParseOperatorPrefix();

            if (CT.Type == TokenType.Operand && CT.Subtype == TokenSubtype.Range)
            {
                var op = CT;
                Next();

                if (_finder != null)
                {
                    var xrange = _finder.Find(_context.Sheet, op.Value);
                    foreach (var cell in xrange.Cells)
                    {
                        cell.Expression = Parse(cell.ExcelFormula, _context);
                    }

                    return xrange.Cells.Length == 1 ? xrange.Cells[0].Expression : XRangeExpression.Create(xrange);
                }

                throw new Exception("ExcelFormulaTokenSubtype is a Range, but no 'CellFinder' class is provided.");
            }

            return left;
        }

        // -
        private Expression ParseOperatorPrefix()
        {
            Expression left = ParseValue();

            while (CT.Type == TokenType.OperatorPrefix && CT.Value == "-")
            {
                var op = CT;
                Next();
                Expression right = ParseValue();

                switch (op.Value)
                {
                    case "-":
                        return Expression.Negate(right);
                    default:
                        throw new NotSupportedException(op.Value);
                }
            }

            return left;
        }

        private Expression ParseValue()
        {
            if (CT.Type == TokenType.Operand)
            {
                if (CT.Subtype == TokenSubtype.Logical)
                {
                    var op = CT;
                    Next();

                    return Expression.Constant(bool.Parse(op.Value));
                }

                if (CT.Subtype == TokenSubtype.Number)
                {
                    var op = CT;
                    Next();

                    return Expression.Constant(double.Parse(op.Value, NumberStyles.Any, CultureInfo.InvariantCulture));
                }

                if (CT.Subtype == TokenSubtype.Text)
                {
                    var op = CT;
                    Next();

                    return Expression.Constant(op.Value);
                }
            }

            return null;
        }

        private Expression Parse(IList<ExcelFormulaToken> tokens, ExcelFormulaContext context)
        {
            if (tokens == null)
            {
                return null;
            }

            return new ExpressionParser(tokens, context, _finder).Parse();
        }
    }
}