using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaParser;
using NFluent;
using Xunit;

namespace ExcelFormulaExpressionParser.Tests
{
    public partial class OperandFunctionTests
    {
        [Fact]
        public void OperandFunction_NestedFunctions0()
        {
            // Assign
            var formula = new ExcelFormula("=ROUND(1.01, 0) + 1");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(2);
        }

        [Fact]
        public void OperandFunction_NestedFunctions1()
        {
            // Assign
            var formula = new ExcelFormula("=ROUND(-10 * 1.001 + 7, 0) + 1");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(-2);
        }

        [Fact]
        public void OperandFunction_NestedFunctions2()
        {
            // Assign
            var formula = new ExcelFormula("=ROUND(-3 * ROUND(1.001 * 2, 1) + 7.001, 1) + 11");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(18);
        }

        [Fact]
        public void OperandFunction_NestedSubexpressions1()
        {
            // Assign
            var formula = new ExcelFormula("=5 * (1 + 2)");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(15);
        }

        [Fact]
        public void OperandFunction_NestedSubexpressions2()
        {
            // Assign
            var formula = new ExcelFormula("=(5 * (1 + 2))");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(15);
        }
    }
}