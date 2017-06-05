using System.Linq.Expressions;
using ExcelFormulaParser;
using NFluent;
using Xunit;

namespace ExcelFormulaExpressionParser.Tests
{
    public class OperandFunctionTests
    {
        [Fact]
        public void OperandFunction_And()
        {
            // Assign
            var formula = new ExcelFormula("=AND(TRUE, TRUE, FALSE)");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = Expression.Lambda(expression).Compile().DynamicInvoke();

            // Assert
            Check.That(result).IsEqualTo(false);
        }

        [Fact]
        public void OperandFunction_If_False()
        {
            // Assign
            var formula = new ExcelFormula("=IF(FALSE, 42, -1)");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = Expression.Lambda(expression).Compile().DynamicInvoke();

            // Assert
            Check.That(result).IsEqualTo(-1);
        }

        [Fact]
        public void OperandFunction_If_True()
        {
            // Assign
            var formula = new ExcelFormula("=IF(TRUE, 42, -1)");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = Expression.Lambda(expression).Compile().DynamicInvoke();

            // Assert
            Check.That(result).IsEqualTo(42);
        }

        [Fact]
        public void OperandFunction_Or()
        {
            // Assign
            var formula = new ExcelFormula("=OR(TRUE, TRUE, FALSE)");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = Expression.Lambda(expression).Compile().DynamicInvoke();

            // Assert
            Check.That(result).IsEqualTo(true);
        }

        [Fact]
        public void OperandFunction_Year()
        {
            // Assign
            var formula = new ExcelFormula("=OR(TRUE, TRUE, FALSE)");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = Expression.Lambda(expression).Compile().DynamicInvoke();

            // Assert
            Check.That(result).IsEqualTo(true);
        }
    }
}