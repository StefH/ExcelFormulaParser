using System;
using System.Collections.Generic;
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
                package.Workbook.Worksheets[1].Cells["B1"].Value = 77;

                var sheets = new List<XSheet>();
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    var sheet = new XSheet(worksheet.Name);

                    // Obtain the worksheet size 
                    ExcelCellAddress startCell = worksheet.Dimension.Start;
                    ExcelCellAddress endCell = worksheet.Dimension.End;

                    for (int r = startCell.Row; r <= endCell.Row; r++)
                    {
                        var xrow = new XRow(sheet);
                        for (int c = startCell.Column; c <= endCell.Column; c++)
                        {
                            xrow.Cells.Add(ToXCell(xrow, worksheet.Cells[r, c]));
                        }

                        sheet.Rows.Add(xrow);
                    }

                    sheets.Add(sheet);
                }

                var calcCell = sheets[0].Rows[3].Cells[1];

                Func<string, string, XCell> findCellBySheetAndAddress = (sheetName, address) =>
                {
                    XSheet sheet;
                    if (!address.Contains('!'))
                    {
                        // Same sheet
                        sheet = sheets.First(s => s.Name == sheetName);
                    }
                    else
                    {
                        // Other sheet
                        string[] parts = address.Split('!');
                        sheet = sheets.First(s => s.Name == parts[0]);
                        address = parts[1];
                    }

                    return sheet.Rows.SelectMany(r => r.Cells).FirstOrDefault(c => c.Address == address);
                };

                var parser = new ExcelFormulaExpressionParser(calcCell.ExcelFormula, (ExcelFormulaContext)calcCell.ExcelFormula.Context, findCellBySheetAndAddress);

                Expression x = parser.Parse();
                System.Console.WriteLine($"Expression = `{x}`");

                var o = x.Optimize();

                System.Console.WriteLine($"Expression Optimize = `{o}`");

                LambdaExpression le = Expression.Lambda(o);

                var result = le.Compile().DynamicInvoke();

                System.Console.WriteLine($"result = `{result}`");

                var bool2 = sheets[0].Rows[4].Cells[1];
                var boolParser = new ExcelFormulaExpressionParser(bool2.ExcelFormula, (ExcelFormulaContext)bool2.ExcelFormula.Context, findCellBySheetAndAddress);

                Expression bx = boolParser.Parse();
                System.Console.WriteLine($"Expression = `{bx}`");

                var o2 = bx.Optimize();

                System.Console.WriteLine($"Expression Optimize = `{o2}`");

                var bresult = Expression.Lambda(o2).Compile().DynamicInvoke();
                System.Console.WriteLine($"bresult = `{bresult}`");

                int u = 0;
            }
        }

        private static XCell ToXCell(XRow row, ExcelRange r)
        {
            var c = new XCell(row)
            {
                Address = r.Address,
                FullAddress = $"{r.Worksheet.Name}!{r.Address}"
            };

            var context = new ExcelFormulaContext
            {
                Sheet = r.Worksheet.Name
            };

            if (r.Value != null)
            {
                c.Value = r.Value;

                if (r.Value is bool)
                {
                    c.ValueFormula = new ExcelFormula((bool) r.Value ? "=TRUE" : "=FALSE", context);
                }
                else if (r.Value.IsNumeric())
                {
                    c.ValueFormula = new ExcelFormula("=" + r.Value, context);
                }
                else
                {
                    c.ValueFormula = new ExcelFormula(string.Format("=\"{0}\"", r.Text), context);
                }
            }

            if (!string.IsNullOrEmpty(r.Formula))
            {
                c.Formula = "=" + r.Formula;
                c.ExcelFormula = new ExcelFormula(c.Formula, context);
            }

            return c;
        }

        private static void CalcTest1()
        {
            // haakjes, machtsverheffen, vermenigvuldigen, delen, worteltrekken, optellen, aftrekken
            var excelFormula = new ExcelFormula("=-(1+2) * ROUND(4/2.7,2) + POWER(1+1,4) + 500 + SIN(3.1415926) + COS(3.1415926/2) + ABS(-1)");
            var parser = new ExcelFormulaExpressionParser(excelFormula);

            Expression x = parser.Parse();
            System.Console.WriteLine($"Expression = `{x}`");

            var o = x.Optimize();

            System.Console.WriteLine($"Expression = `{o}`");

            LambdaExpression le = Expression.Lambda(o);

            var result = le.Compile().DynamicInvoke();

            System.Console.WriteLine($"result = `{result}`");

            var trunc2 = new ExcelFormula("=TRUNC(91.789, 2)");
            System.Console.WriteLine("trunc2 = `{0}`", Expression.Lambda(new ExcelFormulaExpressionParser(trunc2).Parse()).Compile().DynamicInvoke());

            var trunc0 = new ExcelFormula("=TRUNC(91.789)");
            System.Console.WriteLine("trunc0 = `{0}`", Expression.Lambda(new ExcelFormulaExpressionParser(trunc0).Parse()).Compile().DynamicInvoke());
        }
    }
}