using System.Linq.Expressions;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaExpressionParser.Functions;
using NFluent;
using Xunit;

namespace ExcelFormulaExpressionParser.Tests
{
    public class DateFunctionsTests
    {
        [Fact]
        public void Date()
        {
            // Assign

            // Act
            var resultExpression = DateFunctions.Date(Expression.Constant(2017), Expression.Constant(2), Expression.Constant(3));
            var result = resultExpression.LambdaInvoke<double>();

            // Assert
            Check.That(result).IsEqualTo(42769.0);
        }
    }
}
