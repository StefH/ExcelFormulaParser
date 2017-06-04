using ExcelFormulaExpressionParser.Models;

namespace ExcelFormulaExpressionParser.Utils
{
    public static class ExcelUtils
    {
        private static readonly char[] Numbers = "0123456789".ToCharArray();

        public static CellAddress ParseExcelAddress(string value)
        {
            int startIndex = value.IndexOfAny(Numbers);

            return new CellAddress
            {
                Column = ExcelColumnNameToNumber(value.Substring(0, startIndex)),
                Row = int.Parse(value.Substring(startIndex))
            };
        }

        private static int ExcelColumnNameToNumber(string columnName)
        {
            int sum = 0;
            foreach (char c in columnName)
            {
                sum *= 26;
                sum += c - 'A' + 1;
            }

            return sum;
        }
    }
}