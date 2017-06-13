using ExcelFormulaParser;
using NFluent;
using System.Linq.Expressions;
using Xunit;
using ExcelFormulaExpressionParser.Extensions;
using System;
using ExcelFormulaExpressionParser.Helpers;

namespace ExcelFormulaExpressionParser.Tests
{
    public partial class OperandFunctionTests
    {
        [Fact]
        public void OperandFunction_Date_Now()
        {
            // Assign
            var now = DateTimeHelpers.ToOADate(DateTime.Now);
            var formula = new ExcelFormula("=NOW()");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsStrictlyGreaterThan(now);
        }
    }
}