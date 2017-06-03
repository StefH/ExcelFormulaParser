using NFluent;
using Xunit;

namespace ExcelFormulaParser.Tests
{
    public class OperatorMathTests
    {
        // ReSharper disable once InconsistentNaming
        private ExcelFormula ef;

        [Fact]
        public void ExcelFormula_Math_Add()
        {
            // Assign
            const string formula = "=1+";

            // Act
            ef = new ExcelFormula(formula);

            // Assert
            Check.That(ef.Count).Equals(2);

            Check.That(ef[0].Value).Equals("1");
            Check.That(ef[0].Type).Equals(ExcelFormulaTokenType.Operand);
            Check.That(ef[0].Subtype).Equals(ExcelFormulaTokenSubtype.Number);

            Check.That(ef[1].Value).Equals("+");
            Check.That(ef[1].Type).Equals(ExcelFormulaTokenType.OperatorInfix);
            Check.That(ef[1].Subtype).Equals(ExcelFormulaTokenSubtype.Math);
        }

        [Fact]
        public void ExcelFormula_Math_Multiply()
        {
            // Assign
            const string formula = "=1*";

            // Act
            ef = new ExcelFormula(formula);

            // Assert
            Check.That(ef.Count).Equals(2);

            Check.That(ef[0].Value).Equals("1");
            Check.That(ef[0].Type).Equals(ExcelFormulaTokenType.Operand);
            Check.That(ef[0].Subtype).Equals(ExcelFormulaTokenSubtype.Number);

            Check.That(ef[1].Value).Equals("*");
            Check.That(ef[1].Type).Equals(ExcelFormulaTokenType.OperatorInfix);
            Check.That(ef[1].Subtype).Equals(ExcelFormulaTokenSubtype.Math);
        }
    }
}