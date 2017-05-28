using System.Collections.Generic;
using System.Globalization;
using System.Linq.Expressions;

namespace ExcelFormulaParser.Expressions.Console
{
    class ExcelFormulaExpressionParser
    {
        private int _index;

        private ExcelFormulaToken CurrentToken => _list[_index];

        private readonly IList<ExcelFormulaToken> _list;

        public ExcelFormulaExpressionParser(IList<ExcelFormulaToken> list)
        {
            _list = list;
        }

        public Expression Parse()
        {
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
            }

            if (CurrentToken.Type == ExcelFormulaTokenType.Subexpression)
            {
                if (CurrentToken.Subtype == ExcelFormulaTokenSubtype.Start)
                {
                    var list = new List<ExcelFormulaToken>();

                    Next();
                    while (CurrentToken.Subtype != ExcelFormulaTokenSubtype.Stop)
                    {
                        list.Add(CurrentToken);
                        Next();
                    }

                    Next();

                    var subexpressionParser = new ExcelFormulaExpressionParser(list);
                    return subexpressionParser.Parse();
                }
            }

            return null;
        }
    }
}