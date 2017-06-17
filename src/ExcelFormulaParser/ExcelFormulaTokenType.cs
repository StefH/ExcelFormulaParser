namespace ExcelFormulaParser
{
    public enum ExcelFormulaTokenType
    {
        Noop,
        Operand,
        Function,
        Subexpression,
        Argument,
        OperatorPrefix,
        OperatorInfix,
        OperatorPostfix,
        Whitespace,
        Unknown
    }
}