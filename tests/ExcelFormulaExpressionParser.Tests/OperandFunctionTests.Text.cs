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
        public void OperandFunction_Text_Left()
        {
            // Assign
            var formula = new ExcelFormula("=LEFT(\"abc\")");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<string>();

            // Assert
            Check.That(result).IsEqualTo("a");
        }

        [Fact]
        public void OperandFunction_Text_Left_NumChars()
        {
            // Assign
            var formula = new ExcelFormula("=LEFT(\"abc\", 2)");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<string>();

            // Assert
            Check.That(result).IsEqualTo("ab");
        }
    }
}