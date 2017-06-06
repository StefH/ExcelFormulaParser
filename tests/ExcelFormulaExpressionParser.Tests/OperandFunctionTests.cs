using System.Collections.Generic;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Models;
using ExcelFormulaParser;
using Moq;
using NFluent;
using Xunit;

namespace ExcelFormulaExpressionParser.Tests
{
    public class OperandFunctionTests
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

        [Fact]
        public void OperandFunction_VLookup()
        {
            // Assign
            var sheet = new XSheet("sheet1");
            var cells = new List<XCell>();
            var row1 = new XRow(sheet, 1);
            cells.Add(new XCell(row1, "A1") { ExcelFormula = new ExcelFormula("=1")});
            cells.Add(new XCell(row1, "B1") { ExcelFormula = new ExcelFormula("=10") });
            var row2 = new XRow(sheet, 2);
            cells.Add(new XCell(row2, "A2") { ExcelFormula = new ExcelFormula("=2") });
            cells.Add(new XCell(row2, "B2") { ExcelFormula = new ExcelFormula("=20") });

            _finder.Setup(f => f.Find(It.IsAny<string>(), It.IsAny<string>())).Returns(cells);
            var formula = new ExcelFormula("=VLOOKUP(1.5, A1:B2, 2)");

            // Act
            Expression expression = new ExpressionParser(formula, _context, _finder.Object).Parse();
            var result = Expression.Lambda(expression).Compile().DynamicInvoke();

            // Assert
            Check.That(result).IsEqualTo(true);
        }
    }
}