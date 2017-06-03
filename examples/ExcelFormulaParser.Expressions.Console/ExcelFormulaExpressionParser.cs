using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq.Expressions;
using JetBrains.Annotations;

namespace ExcelFormulaParser.Expressions.Console
{
    class ExcelFormulaExpressionParser
    {
        private int _index;

        private ExcelFormulaToken CurrentToken => _list[_index];

        private readonly IList<ExcelFormulaToken> _list;
        private readonly ExcelFormulaContext _context;
        private readonly Func<string, string, XCell> _findCell;

        /// <summary>
        /// ExcelFormulaExpressionParser
        /// </summary>
        /// <param name="list">The ExcelFormula or a list from ExcelFormulaToken.</param>
        /// <param name="context">The ExcelFormulaContext.</param>
        /// <param name="findCell">Function to find a cell by sheetname and address. (optional if no real Excel Workbook is parsed)</param>
        public ExcelFormulaExpressionParser([NotNull] IList<ExcelFormulaToken> list, [CanBeNull] ExcelFormulaContext context = null, [CanBeNull] Func<string, string, XCell> findCell = null)
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
            Expression left = ParseMultiplicative();

            while (CurrentToken.Type == ExcelFormulaTokenType.OperatorInfix && CurrentToken.Subtype == ExcelFormulaTokenSubtype.Math && (CurrentToken.Value == "+" || CurrentToken.Value == "-"))
            {
                var op = CurrentToken;
                Next();
                Expression right = ParseMultiplicative();

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

        // *, /
        private Expression ParseMultiplicative()
        {
            Expression left = ParseOperatorPrefix();

            while (CurrentToken.Type == ExcelFormulaTokenType.OperatorInfix && CurrentToken.Subtype == ExcelFormulaTokenSubtype.Math && (CurrentToken.Value == "*" || CurrentToken.Value == "/"))
            {
                var op = CurrentToken;
                Next();
                Expression right = ParseOperatorPrefix();

                switch (op.Value)
                {
                    case "*":
                        left = Expression.Multiply(left, right);
                        break;
                    case "/":
                        left = Expression.Divide(left, right);
                        break;
                }
            }

            return left;
        }

        // -
        private Expression ParseOperatorPrefix()
        {
            Expression left = ParsePrimary();

            if (CurrentToken.Type == ExcelFormulaTokenType.OperatorPrefix && CurrentToken.Value == "-")
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
            if (CurrentToken.Type == ExcelFormulaTokenType.Operand)
            {
                if (CurrentToken.Subtype == ExcelFormulaTokenSubtype.Number)
                {
                    var op = CurrentToken;
                    Next();

                    return Expression.Constant(double.Parse(op.Value, NumberStyles.Any, CultureInfo.InvariantCulture));
                }

                if (CurrentToken.Subtype == ExcelFormulaTokenSubtype.Text)
                {
                    var op = CurrentToken;
                    Next();

                    return Expression.Constant(op.Value);
                }

                if (CurrentToken.Subtype == ExcelFormulaTokenSubtype.Range)
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

                        var cellExpressionParser = new ExcelFormulaExpressionParser(formula, _context, _findCell);
                        return cellExpressionParser.Parse();
                    }

                    throw new Exception("ExcelFormulaTokenSubtype is a Range, but no 'findCell' function is provided.");
                }
            }

            if (CurrentToken.Type == ExcelFormulaTokenType.Subexpression)
            {
                if (CurrentToken.Subtype == ExcelFormulaTokenSubtype.Start)
                {
                    var tokens = new List<ExcelFormulaToken>();

                    Next();
                    while (CurrentToken.Subtype != ExcelFormulaTokenSubtype.Stop)
                    {
                        tokens.Add(CurrentToken);
                        Next();
                    }

                    Next();

                    var subexpressionParser = new ExcelFormulaExpressionParser(tokens, _context, _findCell);
                    return subexpressionParser.Parse();
                }
            }

            if (CurrentToken.Type == ExcelFormulaTokenType.Function)
            {
                if (CurrentToken.Subtype == ExcelFormulaTokenSubtype.Start)
                {
                    string functionName = CurrentToken.Value;

                    var arguments = new List<Expression>();
                    var tokens = new List<ExcelFormulaToken>();
                    var argumentParser = new ExcelFormulaExpressionParser(tokens, _context, _findCell);

                    Next();
                    while (CurrentToken.Subtype != ExcelFormulaTokenSubtype.Stop)
                    {
                        if (CurrentToken.Type == ExcelFormulaTokenType.Argument)
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

                        case "COS":
                            return MathExpression.Cos(arguments[0]);

                        case "POWER":
                            return Expression.Power(arguments[0], arguments[1]);

                        case "ROUND":
                            return MathExpression.Round(arguments[0], arguments[1]);

                        case "SIN":
                            return MathExpression.Sin(arguments[0]);

                        default:
                            throw new NotImplementedException(functionName);
                    }
                }
            }

            return null;
        }
    }
}