namespace ExcelFormulaParser.Expressions.Console
{
    public class XCell
    {
        public string Address { get; set; }

        public string FullAddress { get; set; }

        public object Value { get; set; }

        public ExcelFormula ValueFormula { get; set; }

        public string Formula { get; set; }

        public ExcelFormula ExcelFormula { get; set; }
    }
}
