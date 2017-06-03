using ExcelFormulaParser;

namespace ExcelFormulaExpressionParser
{
    public class XCell
    {
        public int Column { get; }

        public int Row { get; }

        public string Address { get; }

        public string FullAddress { get; set; }

        public object Value { get; set; }

        // public ExcelFormula ValueFormula { get; set; }

        public string Formula { get; set; }

        public ExcelFormula ExcelFormula { get; set; }

        public XRow XRow { get; }

        public XCell(XRow xRow, string addres)
        {
            XRow = xRow;
            Address = addres;

            var result = ExcelUtils.ParseExcelAddress(addres);
            Column = result.Column;
            Row = result.Row;
        }
    }
}