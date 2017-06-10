﻿using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaParser;
using NFluent;
using Xunit;

namespace ExcelFormulaExpressionParser.Tests
{
    public partial class OperandFunctionTests
    {
        [Fact]
        public void OperandFunction_NestedFunctions1()
        {
            // Assign
            var formula = new ExcelFormula("=ABS(-10 * 1 + 7)");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(-3);
        }

        [Fact]
        public void OperandFunction_NestedFunctions2()
        {
            // Assign
            var formula = new ExcelFormula("=ABS(-10 * ROUND(1.123 * 2) + 7)");

            // Act
            Expression expression = new ExpressionParser(formula).Parse();
            var result = expression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(27);
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