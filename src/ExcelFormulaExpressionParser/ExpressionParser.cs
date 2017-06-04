using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
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
        private readonly Func<string, string, XCell[]> _findCells;

        /// <summary>
        /// ExpressionParser
        /// </summary>
        /// <param name="list">The ExcelFormula or a list from ExcelFormulaToken.</param>
        /// <param name="context">The ExcelFormulaContext.</param>
        /// <param name="findCells">Function to find cells by sheetname and address. (optional if no real Excel Workbook is parsed)</param>
        public ExpressionParser([NotNull] IList<ExcelFormulaToken> list, [CanBeNull] ExcelFormulaContext context = null, [CanBeNull] Func<string, string, XCell[]> findCells = null)
        {
            _list = list;
            _context = context;
            _findCells = findCells;
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
                            left = MathExpression.Power(left, right);
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

                    var subexpressionParser = new ExpressionParser(tokens, _context, _findCells);
                    return subexpressionParser.Parse();
                }
            }

            return left;
        }

        private Expression ParseFunction()
        {
            Expression left = ParseRange();

            if (CT.Type == TokenType.Function)
            {
                if (CT.Subtype == TokenSubtype.Start)
                {
                    string functionName = CT.Value;

                    var arguments = new List<Expression>();
                    var tokens = new List<ExcelFormulaToken>();

                    Next();
                    while (CT.Subtype != TokenSubtype.Stop)
                    {
                        if (CT.Type == TokenType.Argument)
                        {
                            arguments.Add(new ExpressionParser(tokens, _context, _findCells).Parse());

                            tokens.Clear();
                        }
                        else
                        {
                            tokens.Add(CT);
                        }

                        Next();
                    }

                    arguments.Add(new ExpressionParser(tokens, _context, _findCells).Parse());

                    Next();

                    switch (functionName)
                    {
                        case "ABS": return MathExpression.Abs(arguments[0]);
                        case "AND": return LogicalExpression.And(arguments);
                        case "COS": return MathExpression.Cos(arguments[0]);
                        case "IF": return Expression.Condition(arguments[0], arguments[1], arguments[2]);
                        case "OR": return LogicalExpression.Or(arguments);
                        case "POWER": return MathExpression.Power(arguments[0], arguments[1]);
                        case "ROUND": return MathExpression.Round(arguments[0], arguments[1]);
                        case "SIN": return MathExpression.Sin(arguments[0]);
                        case "SUM": return MathExpression.Sum(arguments);
                        case "TRUNC": return MathExpression.Trunc(arguments);

                        default:
                            throw new NotImplementedException(functionName);
                    }
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

                if (_findCells != null)
                {
                    var cells = _findCells(_context.Sheet, op.Value); // B1 or 'Sheet1'!B1

                    if (cells.Length == 0)
                    {
                        throw new Exception("No cell(s) found.");
                    }

                    ExcelFormula excel;
                    if (cells.Length == 1)
                    {
                        var cell = cells[0];
                        excel = cell.ExcelFormula;
                    }
                    else
                    {
                        //var argToken = ExcelFormula.CreateArgumentToken();
                        //tokens = cells.Select(c => c.ExcelFormula.ToList()).Aggregate((total, next) =>
                        //{
                        //    var l = new List<ExcelFormulaToken>();
                        //    l.AddRange(total);
                        //    l.Add(argToken);
                        //    l.AddRange(next);
                        //    return l;
                        //});

                        //tokens = ExcelFormula.WrapInSubExpression(tokens);
                        //int y = 9;

                        string form = string.Join(", ", cells.Select(c => c.Value).Where(v => v != null));
                        excel = new ExcelFormula($"=({form})");
                    }

                    var cellExpressionParser = new ExpressionParser(excel, _context, _findCells);
                    return cellExpressionParser.Parse();
                }

                throw new Exception("ExcelFormulaTokenSubtype is a Range, but no 'findCells' function is provided.");

            }

            return null;
        }
    }
}