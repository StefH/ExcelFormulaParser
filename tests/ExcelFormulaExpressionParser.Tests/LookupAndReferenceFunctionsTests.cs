using System.Collections.Generic;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaExpressionParser.Functions;
using ExcelFormulaExpressionParser.Models;
using ExcelFormulaParser;
using Microsoft.VisualStudio.TestPlatform.CommunicationUtilities.EventHandlers;
using NFluent;
using Xunit;
using Xunit.Sdk;

namespace ExcelFormulaExpressionParser.Tests
{
    public class LookupAndReferenceFunctionsTests
    {
        private readonly XRange _range;

        public LookupAndReferenceFunctionsTests()
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

            _range = new XRange
            {
                Cells = cells.ToArray(),
                Address = "A10:B12",
                Sheet = sheet,
                Start = new CellAddress { Row = 10, Column = 1 },
                End = new CellAddress { Row = 12, Column = 2 }
            };
        }

        [Fact]
        public void VLookup_Approximate_With_FALSE()
        {
            // Assign

            // Act
            var resultExpression = LookupAndReferenceFunctions.VLookup(Expression.Constant(2.1), _range, Expression.Constant(2), Expression.Constant(true));
            var result = resultExpression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(20.0);
        }

        [Fact]
        public void VLookup_Approximate_With_1()
        {
            // Assign

            // Act
            var resultExpression = LookupAndReferenceFunctions.VLookup(Expression.Constant(2.1), _range, Expression.Constant(2), Expression.Constant(1));
            var result = resultExpression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(20.0);
        }

        [Fact]
        public void VLookup_ExactMatch_Found()
        {
            // Assign

            // Act
            var resultExpression = LookupAndReferenceFunctions.VLookup(Expression.Constant(1.0), _range, Expression.Constant(2), Expression.Constant(false));
            var result = resultExpression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(10.0);
        }

        [Fact]
        public void VLookup_ExactMatch_NotFound()
        {
            // Assign

            // Act
            var resultExpression = LookupAndReferenceFunctions.VLookup(Expression.Constant(1.1), _range, Expression.Constant(2), Expression.Constant(false));
            var result = resultExpression.LambdaInvoke<string>();

            // Assert
            Check.That(result).IsEqualTo("#N/A");
        }
    }
}