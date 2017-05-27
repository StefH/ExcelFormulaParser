namespace ExcelFormulaParser
{
    public class ExcelFormulaToken
    {
        public string Value { get; internal set; }

        public ExcelFormulaTokenType Type { get; internal set; }

        public ExcelFormulaTokenSubtype Subtype { get; internal set; }

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