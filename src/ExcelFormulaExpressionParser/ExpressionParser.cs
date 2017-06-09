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

            return ParseAdditive();
        }

        private void Next()
        {
            if (_index + 1 < _list.Count)
            {
                _index++;
            }
        }

        // +, -
        private Expression ParseAdditive()
        {
            Expression left = ParseMultiplication();

            if (CT.Type == TokenType.OperatorInfix && CT.Subtype == TokenSubtype.Math && (CT.Value == "+" || CT.Value == "-"))
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

            if (CT.Type == TokenType.OperatorInfix)
            {
                if (CT.Subtype == TokenSubtype.Math && (CT.Value == "^" || CT.Value == "*" || CT.Value == "/"))
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
            }

            return left;
        }

        // >, >=, <, <=, <>, =
        private Expression ParseLogical()
        {
            Expression left = ParseOperatorPrefix();

            if (CT.Type == TokenType.OperatorInfix && CT.Subtype == TokenSubtype.Logical)
            {
                var op = CT;
                Next();
                Expression right = ParseOperatorPrefix();

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

        // -
        private Expression ParseOperatorPrefix()
        {
            Expression left = ParseValue();

            if (CT.Type == TokenType.OperatorPrefix && CT.Value == "-")
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
            Expression left = ParseSubexpression();

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

            return left;
        }

        private Expression ParseSubexpression()
        {
            Expression left = ParseFunction();

            if (CT.Type == TokenType.Subexpression)
            {
                if (CT.Subtype == TokenSubtype.Start)
                {
                    var tokens = new List<ExcelFormulaToken>();

                    Next();
                    while (CT.Subtype != TokenSubtype.Stop)
                    {
                        tokens.Add(CT);
                        Next();
                    }

                    Next();

                    return Parse(tokens, _context);
                }
            }

            return left;
        }

        private Expression ParseFunction()
        {
            Expression left = ParseRange();

            if (CT.Type == TokenType.Function && CT.Subtype == TokenSubtype.Start)
            {
                void AddToArgumentsList(ICollection<object> args, Expression expression)
                {
                    var constantExpression = expression as ConstantExpression;
                    if (constantExpression != null && constantExpression.Type == typeof(XRange))
                    {
                        args.Add(constantExpression.Value);
                    }
                    else
                    {
                        args.Add(expression);
                    }
                }

                string functionName = CT.Value;

                var arguments = new List<object>();
                var tokens = new List<ExcelFormulaToken>();

                Next();

                while (CT.Subtype != TokenSubtype.Stop)
                {
                    if (CT.Type == TokenType.Function && CT.Subtype == TokenSubtype.Start)
                    {
                        // Nested function
                        AddToArgumentsList(arguments, ParseFunction());
                    }
                    else if (CT.Type == TokenType.Argument)
                    {
                        AddToArgumentsList(arguments, Parse(tokens, _context));

                        tokens.Clear();
                    }
                    else
                    {
                        tokens.Add(CT);
                    }

                    Next();
                }

                if (tokens.Any())
                {
                    AddToArgumentsList(arguments, Parse(tokens, _context));
                }

                Next();

                var expressions = arguments.Where(a => a is Expression).Cast<Expression>().ToArray();
                var xranges = arguments.Where(a => a is XRange).Cast<XRange>().ToArray();

                switch (functionName)
                {
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
                    case "ROUND": return MathFunctions.Round(expressions[0], expressions[1]);
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

            return left;
        }

        private Expression ParseRange()
        {
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

                    return xrange.Cells.Length == 1 ? xrange.Cells[0].Expression : Expression.Constant(xrange);
                }

                throw new Exception("ExcelFormulaTokenSubtype is a Range, but no 'CellFinder' class is provided.");
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