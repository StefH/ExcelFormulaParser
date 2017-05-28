using NFluent;
using Xunit;

namespace ExcelFormulaParser.Tests
{
    public class ExcelFormulaTests
    {
        // ReSharper disable once InconsistentNaming
        private ExcelFormula ef;

        [Fact]
        public void ExcelFormula_ParseNumber()
        {
            // Assign
            const string formula = "=42";

            // Act
            ef = new ExcelFormula(formula);

            // Assert
            Check.That(ef.Count).Equals(1);
            Check.That(ef[0].Value).Equals("42");
            Check.That(ef[0].Type).Equals(ExcelFormulaTokenType.Operand);
            Check.That(ef[0].Subtype).Equals(ExcelFormulaTokenSubtype.Number);
        }

        [Fact]
        public void ExcelFormula_ParseAddIntegerNumbers()
        {
            // Assign
            const string formula = "=1+2";

            // Act
            ef = new ExcelFormula(formula);

            // Assert
            Check.That(ef.Count).Equals(3);
            Check.That(ef[0].Value).Equals("1");
            Check.That(ef[0].Type).Equals(ExcelFormulaTokenType.Operand);
            Check.That(ef[0].Subtype).Equals(ExcelFormulaTokenSubtype.Number);

            Check.That(ef[1].Value).Equals("+");
            Check.That(ef[1].Type).Equals(ExcelFormulaTokenType.OperatorInfix);
            Check.That(ef[1].Subtype).Equals(ExcelFormulaTokenSubtype.Math);

            Check.That(ef[2].Value).Equals("2");
            Check.That(ef[2].Type).Equals(ExcelFormulaTokenType.Operand);
            Check.That(ef[2].Subtype).Equals(ExcelFormulaTokenSubtype.Number);
        }

        [Fact]
        public void ExcelFormula_ParseDoubleNumbers()
        {
            // Assign
            const string formula = "=1.1+2.2";

            // Act
            ef = new ExcelFormula(formula);

            // Assert
            Check.That(ef.Count).Equals(3);
            Check.That(ef[0].Value).Equals("1.1");
            Check.That(ef[2].Value).Equals("2.2");
        }
    }
}