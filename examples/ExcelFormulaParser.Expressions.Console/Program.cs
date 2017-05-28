using System.Linq.Expressions;

namespace ExcelFormulaParser.Expressions.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            // haakjes, machtsverheffen, vermenigvuldigen, delen, worteltrekken, optellen, aftrekken
            var excelFormula = new ExcelFormula("=-(1+2) * 4 / 2.7");
            var parser = new ExcelFormulaExpressionParser(excelFormula);

            Expression x = parser.Parse();

            LambdaExpression le = Expression.Lambda(x);

            var result = le.Compile().DynamicInvoke();

            System.Console.WriteLine($"result = `{result}`");
        }
    }
}