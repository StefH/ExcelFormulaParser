namespace ExcelFormulaExpressionParser.Models
{
    interface IArgument<T>
    {
        T Value { get; set; }
    }
}
