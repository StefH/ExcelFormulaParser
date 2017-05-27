using NFluent;
using Xunit;

namespace ExcelFormulaParser.Tests
{
    public class ExcelFormulaTests
    {
        // ReSharper disable once InconsistentNaming
        private ExcelFormula excelFormula;

        [Fact]
        public void ExcelFormula_ParseNumber()
        {
            // Assign
            const string formula = "=42";

            // Act
            excelFormula = new ExcelFormula(formula);

            // Assert
            Check.That(excelFormula.Count).Equals(1);
            Check.That(excelFormula[0].Value).Equals("42");
            Check.That(excelFormula[0].Type).Equals(ExcelFormulaTokenType.Operand);
            Check.That(excelFormula[0].Subtype).Equals(ExcelFormulaTokenSubtype.Number);
        }
    }
}