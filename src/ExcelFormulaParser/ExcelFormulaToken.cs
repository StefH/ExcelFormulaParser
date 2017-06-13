namespace ExcelFormulaParser
{
    public class ExcelFormulaToken
    {
        public string Value { get; set; }

        public ExcelFormulaTokenType Type { get; set; }

        public ExcelFormulaTokenSubtype Subtype { get; set; }

        internal ExcelFormulaToken(string value, ExcelFormulaTokenType type) : this(value, type, ExcelFormulaTokenSubtype.Nothing)
        {
        }

        internal ExcelFormulaToken(string value, ExcelFormulaTokenType type, ExcelFormulaTokenSubtype subtype)
        {
            Value = value;
            Type = type;
            Subtype = subtype;
        }
    }
}