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
        public void OperandFunction_NestedFunctions()
        {
            // Assign
            var formula = new ExcelFormula("=SUM(ROUND(1.123, 0),3)");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(4);
        }

        [Fact]
        public void OperandFunction_NestedSubexpressions()
        {
            // Assign
            var formula = new ExcelFormula("=(1 + (2))");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(3);
        }
    }
}