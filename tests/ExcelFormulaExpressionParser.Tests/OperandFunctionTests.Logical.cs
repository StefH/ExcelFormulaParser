using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaParser;
using Moq;
using NFluent;
using Xunit;

namespace ExcelFormulaExpressionParser.Tests
{
    public partial class OperandFunctionTests
    {
        private readonly ExcelFormulaContext _context;
        private readonly Mock<ICellFinder> _finder;

        public OperandFunctionTests()
        {
            _context = new ExcelFormulaContext
            {
                Sheet = "sheet1"
            };

            _finder = new Mock<ICellFinder>();
        }

        [Fact]
        public void OperandFunction_And()
        {
            // Assign
            var formula = new ExcelFormula("=AND(TRUE, FALSE, TRUE, TRUE, FALSE, FALSE)");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<bool>();

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
            var result = expression.LambdaInvoke<double>();

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
            var result = expression.LambdaInvoke<double>();

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
    }
}