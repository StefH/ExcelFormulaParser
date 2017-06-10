using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaParser;
using NFluent;
using Xunit;

namespace ExcelFormulaExpressionParser.Tests
{
    partial class OperandFunctionTests
    {
        [Fact]
        public void OperandFunction_Math_Multiply()
        {
            // Assign
            var formula = new ExcelFormula("=3 * 7 * 9");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(189);
        }

        [Fact]
        public void OperandFunction_Math_Multiply_Negative()
        {
            // Assign
            var formula = new ExcelFormula("=-3 * 7");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(-21);
        }

        [Fact]
        public void OperandFunction_Math_Multiply_Negative_Both()
        {
            // Assign
            var formula = new ExcelFormula("=-3 * -7");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(21);
        }

        [Fact]
        public void OperandFunction_Math_Order()
        {
            // Assign
            var formula = new ExcelFormula("=10 * 50 + 20");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(520);
        }
    }
}