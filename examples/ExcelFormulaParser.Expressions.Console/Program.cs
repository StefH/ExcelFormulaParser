using System;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace ExcelFormulaParser.Expressions.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelTest();
            CalcTest1();
        }

        private static void ExcelTest()
        {
            using (var package = new ExcelPackage(new FileInfo("Book.xlsx")))
            {
                package.Workbook.Worksheets.First().Cells["B1"].Value = 77;

                var row = new XRow();
                row.Cells.Add(ToXCell(package.Workbook.Worksheets.First().Cells["A1"]));
                row.Cells.Add(ToXCell(package.Workbook.Worksheets.First().Cells["B1"]));
                row.Cells.Add(ToXCell(package.Workbook.Worksheets.First().Cells["B2"]));
                row.Cells.Add(ToXCell(package.Workbook.Worksheets.First().Cells["B3"]));

                var cell4 = ToXCell(package.Workbook.Worksheets.First().Cells["B4"]);
                row.Cells.Add(cell4);

                row.Cells.Add(ToXCell(package.Workbook.Worksheets[2].Cells["A1"]));

                Func<string, XCell> find = address =>
                {
                    return row.Cells.FirstOrDefault(c => address.Contains('!') ? c.FullAddress == address : c.Address == address);
                };

                var parser = new ExcelFormulaExpressionParser(cell4.ExcelFormula, find);

                Expression x = parser.Parse();
                System.Console.WriteLine($"Expression = `{x}`");

                var o = x.Optimize();

                System.Console.WriteLine($"Expression = `{o}`");

                LambdaExpression le = Expression.Lambda(o);

                var result = le.Compile().DynamicInvoke();

                System.Console.WriteLine($"result = `{result}`");

                int u = 0;
            }
        }

        private static XCell ToXCell(ExcelRange r)
        {
            var c = new XCell
            {
                Address = r.Address,
                FullAddress = $"{r.Worksheet.Name}!{r.Address}"
            };

            if (r.Value != null)
            {
                c.Value = r.Value;

                if (r.Value.IsNumeric())
                {
                    c.ValueFormula = new ExcelFormula("=" + r.Value);
                }
                else
                {
                    c.ValueFormula = new ExcelFormula(string.Format("=\"{0}\"", r.Text));
                }
            }

            if (!string.IsNullOrEmpty(r.Formula))
            {
                c.Formula = "=" + r.Formula;
                c.ExcelFormula = new ExcelFormula(c.Formula);
            }

            return c;
        }

        private static void CalcTest1()
        {
            // haakjes, machtsverheffen, vermenigvuldigen, delen, worteltrekken, optellen, aftrekken
            var excelFormula = new ExcelFormula("=-(1+2) * ROUND(4 / 2.7, 2) + POWER(1+1,4) + 500 + SIN(3.1415926) + COS(3.1415926 / 2) + ABS(-1)");
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