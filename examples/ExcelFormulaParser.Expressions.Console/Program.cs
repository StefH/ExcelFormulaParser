using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using OfficeOpenXml;

namespace ExcelFormulaParser.Expressions.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelTest();
            CalcTest();
        }

        private static void ExcelTest()
        {
            FileInfo f = new FileInfo("Book.xlsx");
            using (var package = new ExcelPackage(f))
            {
                var row = new XRow();
                row.Cells.Add(ToXCell(package.Workbook.Worksheets.First().Cells["B1"]));
                row.Cells.Add(ToXCell(package.Workbook.Worksheets.First().Cells["B2"]));
                row.Cells.Add(ToXCell(package.Workbook.Worksheets.First().Cells["B3"]));

                int u = 0;
            }
        }

        private static XCell ToXCell(ExcelRange r)
        {
            var c = new XCell
            {
                Address = r.Address,
                Value = r.Value
            };

            if (!string.IsNullOrEmpty(r.Formula))
            {
                c.Formula = "=" + r.Formula;
                c.ExcelFormula = new ExcelFormula(c.Formula);
            }

            return c;
        }

        private static void CalcTest()
        {
            // haakjes, machtsverheffen, vermenigvuldigen, delen, worteltrekken, optellen, aftrekken
            var excelFormula = new ExcelFormula("=-(1+2) * ROUND(4 / 2.7, 2) + POWER(1+1,4) + 500 + SIN(3.1415926)");
            var parser = new ExcelFormulaExpressionParser(excelFormula);

            Expression x = parser.Parse();
            System.Console.WriteLine($"Expression = `{x}`");

            var o = x.Optimize();

            System.Console.WriteLine($"Expression = `{o}`");

            LambdaExpression le = Expression.Lambda(o);

            var result = le.Compile().DynamicInvoke();

            System.Console.WriteLine($"result = `{result}`");
        }
    }
}