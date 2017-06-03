﻿using System;
using System.Collections.Generic;
using System.Globalization;
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

        private ExcelFormulaToken CurrentToken => _list[_index];

        private readonly IList<ExcelFormulaToken> _list;
        private readonly ExcelFormulaContext _context;
        private readonly Func<string, string, XCell> _findCell;

        /// <summary>
        /// ExpressionParser
        /// </summary>
        /// <param name="list">The ExcelFormula or a list from ExcelFormulaToken.</param>
        /// <param name="context">The ExcelFormulaContext.</param>
        /// <param name="findCell">Function to find a cell by sheetname and address. (optional if no real Excel Workbook is parsed)</param>
        public ExpressionParser([NotNull] IList<ExcelFormulaToken> list, [CanBeNull] ExcelFormulaContext context = null, [CanBeNull] Func<string, string, XCell> findCell = null)
        {
            _list = list;
            _context = context;
            _findCell = findCell;
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
            Expression left = ParseOperatorInfix();

            if (CurrentToken.Type == TokenType.OperatorInfix && CurrentToken.Subtype == TokenSubtype.Math && (CurrentToken.Value == "+" || CurrentToken.Value == "-"))
            {
                var op = CurrentToken;
                Next();
                Expression right = ParseOperatorInfix();

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
        private Expression ParseOperatorInfix()
        {
            Expression left = ParseOperatorPrefix();

            if (CurrentToken.Type == TokenType.OperatorInfix)
            {
                if (CurrentToken.Subtype == TokenSubtype.Math && (CurrentToken.Value == "^" || CurrentToken.Value == "*" || CurrentToken.Value == "/"))
                {
                    var op = CurrentToken;
                    Next();
                    Expression right = ParseOperatorPrefix();

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
                else if (CurrentToken.Subtype == TokenSubtype.Logical)
                {
                    var op = CurrentToken;
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
            }

            return left;
        }

        // -
        private Expression ParseOperatorPrefix()
        {
            Expression left = ParsePrimary();

            if (CurrentToken.Type == TokenType.OperatorPrefix && CurrentToken.Value == "-")
            {
                var op = CurrentToken;
                Next();
                Expression right = ParsePrimary();

                switch (op.Value)
                {
                    case "-":
                        return Expression.Negate(right);
                }
            }

            return left;
        }

        private Expression ParsePrimary()
        {
            if (CurrentToken.Type == TokenType.Operand)
            {
                if (CurrentToken.Subtype == TokenSubtype.Logical)
                {
                    var op = CurrentToken;
                    Next();

                    return Expression.Constant(bool.Parse(op.Value));
                }

                if (CurrentToken.Subtype == TokenSubtype.Number)
                {
                    var op = CurrentToken;
                    Next();

                    return Expression.Constant(double.Parse(op.Value, NumberStyles.Any, CultureInfo.InvariantCulture));
                }

                if (CurrentToken.Subtype == TokenSubtype.Text)
                {
                    var op = CurrentToken;
                    Next();

                    return Expression.Constant(op.Value);
                }

                if (CurrentToken.Subtype == TokenSubtype.Range)
                {
                    var op = CurrentToken;
                    Next();

                    if (_findCell != null)
                    {
                        var cell = _findCell(_context.Sheet, op.Value); // B1 or 'Sheet1'!B1

                        ExcelFormula formula;
                        if (cell.ExcelFormula != null)
                        {
                            formula = cell.ExcelFormula;
                        }
                        else if (cell.ValueFormula != null)
                        {
                            formula = cell.ValueFormula;
                        }
                        else
                        {
                            formula = new ExcelFormula("=");
                        }

                        var cellExpressionParser = new ExpressionParser(formula, _context, _findCell);
                        return cellExpressionParser.Parse();
                    }

                    throw new Exception("ExcelFormulaTokenSubtype is a Range, but no 'findCell' function is provided.");
                }
            }

            if (CurrentToken.Type == TokenType.Subexpression)
            {
                if (CurrentToken.Subtype == TokenSubtype.Start)
                {
                    var tokens = new List<ExcelFormulaToken>();

                    Next();
                    while (CurrentToken.Subtype != TokenSubtype.Stop)
                    {
                        tokens.Add(CurrentToken);
                        Next();
                    }

                    Next();

                    var subexpressionParser = new ExpressionParser(tokens, _context, _findCell);
                    return subexpressionParser.Parse();
                }
            }

            if (CurrentToken.Type == TokenType.Function)
            {
                if (CurrentToken.Subtype == TokenSubtype.Start)
                {
                    string functionName = CurrentToken.Value;

                    var arguments = new List<Expression>();
                    var tokens = new List<ExcelFormulaToken>();
                    var argumentParser = new ExpressionParser(tokens, _context, _findCell);

                    Next();
                    while (CurrentToken.Subtype != TokenSubtype.Stop)
                    {
                        if (CurrentToken.Type == TokenType.Argument)
                        {
                            arguments.Add(argumentParser.Parse());

                            tokens.Clear();
                        }
                        else
                        {
                            tokens.Add(CurrentToken);
                        }

                        Next();
                    }

                    arguments.Add(argumentParser.Parse());

                    Next();

                    switch (functionName)
                    {
                        case "ABS":
                            return MathExpression.Abs(arguments[0]);

                        case "AND":
                            return LogicalExpression.And(arguments);

                        case "COS":
                            return MathExpression.Cos(arguments[0]);

                        case "IF":
                            return Expression.Condition(arguments[0], arguments[1], arguments[2]);

                        case "OR":
                            return LogicalExpression.Or(arguments);

                        case "POWER":
                            return MathExpression.Power(arguments[0], arguments[1]);

                        case "ROUND":
                            return MathExpression.Round(arguments[0], arguments[1]);

                        case "SIN":
                            return MathExpression.Sin(arguments[0]);

                        case "TRUNC":
                            return MathExpression.Trunc(arguments);

                        default:
                            throw new NotImplementedException(functionName);
                    }
                }
            }

            return null;
        }
    }
}