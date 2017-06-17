using System.Collections.Generic;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaExpressionParser.Models;
using ExcelFormulaParser;
using Moq;
using NFluent;
using Xunit;

namespace ExcelFormulaExpressionParser.Tests
{
    public partial class OperandFunctionTests
    {
        [Fact]
        public void OperandFunction_VLookup()
        {
            // Assign
            var range = CreateXRange();

            _finder.Setup(f => f.Find(It.IsAny<XSheet>(), "A10:B12")).Returns(range);
            var formula = new ExcelFormula("=VLOOKUP(2.1, A10:B12, 2)");

            // Act
            Expression expression = new ExpressionParser(formula, 0, _context, _finder.Object).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(20.0);
        }

        [Fact]
        public void OperandFunction_VLookup_ExactMatch_Found()
        {
            // Assign
            var range = CreateXRange();

            _finder.Setup(f => f.Find(It.IsAny<XSheet>(), "A10:B12")).Returns(range);
            var formula = new ExcelFormula("=VLOOKUP(1.0, A10:B12, 2, FALSE)");

            // Act
            Expression expression = new ExpressionParser(formula, 0, _context, _finder.Object).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(10.0);
        }

        private static XRange CreateXRange()
        {
            var sheet = new XSheet("sheet1");
            var cells = new List<XCell>();
            var row1 = new XRow(sheet, 1);
            cells.Add(new XCell(row1, "A10") { ExcelFormula = new ExcelFormula("=1"), Expression = Expression.Constant(1.0) });
            cells.Add(new XCell(row1, "B10") { ExcelFormula = new ExcelFormula("=SUM(1,9)"), Expression = Expression.Add(Expression.Constant(1.0), Expression.Constant(9.0)) });
            var row2 = new XRow(sheet, 2);
            cells.Add(new XCell(row2, "A11") { ExcelFormula = new ExcelFormula("=2"), Expression = Expression.Constant(2.0) });
            cells.Add(new XCell(row2, "B11") { ExcelFormula = new ExcelFormula("=20"), Expression = Expression.Constant(20.0) });
            var row3 = new XRow(sheet, 3);
            cells.Add(new XCell(row3, "A12") { ExcelFormula = new ExcelFormula("=3"), Expression = Expression.Constant(3.0) });
            cells.Add(new XCell(row3, "B12") { ExcelFormula = null, Expression = null });

            return new XRange
            {
                Cells = cells.ToArray(),
                Address = "A10:B12",
                Sheet = sheet,
                Start = new CellAddress { Row = 10, Column = 1 },
                End = new CellAddress { Row = 12, Column = 2 }
            };
        }
    }
}