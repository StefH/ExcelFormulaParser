namespace ExcelFormulaParser
{
    public enum ExcelFormulaTokenType
    {
        Noop,
        Operand,
        Function,
        SubExpression,
        Argument,
        OperatorPrefix,
        OperatorInfix,
        OperatorPostfix,
        Whitespace,
        Unknown
    }
}