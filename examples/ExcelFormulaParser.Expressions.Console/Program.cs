using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelFormulaParser.Expressions.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            var excelFormula = new ExcelFormula("=1+2");

            Expression main = null;

            Stack<Expression> stack = new Stack<Expression>();
            foreach (ExcelFormulaToken token in excelFormula)
            {
                stack.Push(HandleType(token));
            }

            int end = 0;
        }

        private static Expression HandleType(ExcelFormulaToken token)
        {
            if (token.Type == ExcelFormulaTokenType.Operand)
            {
                if (token.Subtype == ExcelFormulaTokenSubtype.Number)
                {
                    return Expression.Constant(double.Parse(token.Value));
                }
            }

            return null;
        }
    }
}