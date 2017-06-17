using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaParser;
using NFluent;
using Xunit;

namespace ExcelFormulaExpressionParser.Tests
{
    public class OperatorPrefixLogicalTests
    {
        [Fact]
        public void OperatorPrefixLogical_Equal()
        {
            // Assign
            var formula = new ExcelFormula("=1=2");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<bool>();

            // Assert
            Check.That(result).IsEqualTo(false);
        }

        [Fact]
        public void OperatorPrefixLogical_Gt()
        {
            // Assign
            var formula = new ExcelFormula("=1>2");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<bool>();

            // Assert
            Check.That(result).IsEqualTo(false);
        }

        [Fact]
        public void OperatorPrefixLogical_Gte()
        {
            // Assign
            var formula = new ExcelFormula("=1>=1");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<bool>();

            // Assert
            Check.That(result).IsEqualTo(true);
        }

        [Fact]
        public void OperatorPrefixLogical_Lt()
        {
            // Assign
            var formula = new ExcelFormula("=1<2");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<bool>();

            // Assert
            Check.That(result).IsEqualTo(true);
        }

        [Fact]
        public void OperatorPrefixLogical_Lte()
        {
            // Assign
            var formula = new ExcelFormula("=1<=1");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<bool>();

            // Assert
            Check.That(result).IsEqualTo(true);
        }

        [Fact]
        public void OperatorPrefixLogical_NotEqual()
        {
            // Assign
            var formula = new ExcelFormula("=1<>2");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<bool>();

            // Assert
            Check.That(result).IsEqualTo(true);
        }
    }
}