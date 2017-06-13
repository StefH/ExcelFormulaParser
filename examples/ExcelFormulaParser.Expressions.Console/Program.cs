using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaExpressionParser.Helpers;
using ExcelFormulaExpressionParser.Models;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Utilities;
using System.Globalization;

namespace ExcelFormulaParser.Expressions.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            //Test();
            //CalcTest();
            ExcelTest();
        }

        private static void Test()
        {
            using (var package = new ExcelPackage(new FileInfo("c:\\temp\\test.xlsx")))
            {
                var sheets = new List<XSheet>();
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    var sheet = new XSheet(worksheet.Name);

                    // Obtain the worksheet size 
                    ExcelCellAddress startCell = worksheet.Dimension.Start;
                    ExcelCellAddress endCell = worksheet.Dimension.End;

                    for (int r = startCell.Row; r <= endCell.Row; r++)
                    {
                        var xrow = new XRow(sheet, r);
                        for (int c = startCell.Column; c <= endCell.Column; c++)
                        {
                            xrow.Cells.Add(ToXCell(xrow, worksheet.Cells[r, c]));
                        }

                        sheet.Rows.Add(xrow);
                    }

                    sheets.Add(sheet);
                }

                var calcCell = sheets[2].Rows[16].Cells[2];
                var parser = new ExpressionParser(calcCell.ExcelFormula, (ExcelFormulaContext)calcCell.ExcelFormula.Context, sheets);

                Expression x = parser.Parse();
                System.Console.WriteLine($"Expression = `{x}`");

                var o = x.Optimize();

                System.Console.WriteLine($"Expression Optimize = `{o}`");

                var result = o.LambdaInvoke<double>();

                System.Console.WriteLine($"result = `{result}`");
            }
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
                        var xrow = new XRow(sheet, r);
                        for (int c = startCell.Column; c <= endCell.Column; c++)
                        {
                            xrow.Cells.Add(ToXCell(xrow, worksheet.Cells[r, c]));
                        }

                        sheet.Rows.Add(xrow);
                    }

                    sheets.Add(sheet);
                }

                var calcCell = sheets[0].Rows[3].Cells[1];
                var parser = new ExpressionParser(calcCell.ExcelFormula, (ExcelFormulaContext)calcCell.ExcelFormula.Context, sheets);

                Expression x = parser.Parse();
                System.Console.WriteLine($"Expression = `{x}`");

                var o = x.Optimize();

                System.Console.WriteLine($"Expression Optimize = `{o}`");

                LambdaExpression le = Expression.Lambda(o);

                var result = le.Compile().DynamicInvoke();

                System.Console.WriteLine($"result = `{result}`");

                var bool2 = sheets[0].Rows[4].Cells[1];
                var boolParser = new ExpressionParser(bool2.ExcelFormula, (ExcelFormulaContext)bool2.ExcelFormula.Context, sheets);

                Expression bx = boolParser.Parse();
                System.Console.WriteLine($"Expression = `{bx}`");

                var o2 = bx.Optimize();

                System.Console.WriteLine($"Expression Optimize = `{o2}`");

                var bresult = Expression.Lambda(o2).Compile().DynamicInvoke();
                System.Console.WriteLine($"bresult = `{bresult}`");


                var sum = sheets[0].Rows[6].Cells[1];
                var sumParser = new ExpressionParser(sum.ExcelFormula, (ExcelFormulaContext)sum.ExcelFormula.Context, sheets);

                Expression sumE = sumParser.Parse();
                System.Console.WriteLine($"Expression = `{sumE}`");

                var sumEOpt = sumE.Optimize();

                System.Console.WriteLine($"Expression Optimize = `{sumEOpt}`");

                var sumresult = Expression.Lambda(sumEOpt).Compile().DynamicInvoke();
                System.Console.WriteLine($"sumresult = `{sumresult}`");


                var sumSqrt = sheets[3].Rows[0].Cells[1];
                var sumSqrtParser = new ExpressionParser(sumSqrt.ExcelFormula, (ExcelFormulaContext)sumSqrt.ExcelFormula.Context, sheets);

                Expression sumSqrtE = sumSqrtParser.Parse();
                //System.Console.WriteLine($"Expression = `{sumSqrtE}`");

                var sumSqrtEOpt = sumSqrtE.Optimize();

                //System.Console.WriteLine($"Expression Optimize = `{sumSqrtEOpt}`");

                var sumSqrtresult = Expression.Lambda(sumSqrtEOpt).Compile().DynamicInvoke();
                System.Console.WriteLine($"sumresult = `{sumSqrtresult}`");


                var dateyear = sheets[0].Rows[7].Cells[2];
                var dateyearParser = new ExpressionParser(dateyear.ExcelFormula, (ExcelFormulaContext)dateyear.ExcelFormula.Context, sheets);

                Expression dateyearE = dateyearParser.Parse();
                //System.Console.WriteLine($"Expression = `{sumSqrtE}`");

                var dateyearO = dateyearE.Optimize();

                //System.Console.WriteLine($"Expression Optimize = `{sumSqrtEOpt}`");

                var dateyearResult = Expression.Lambda(dateyearO).Compile().DynamicInvoke();
                System.Console.WriteLine($"dateyearResult = `{dateyearResult}`");

                var dateyearnow = sheets[0].Rows[8].Cells[2];
                var dateyearnowParser = new ExpressionParser(dateyearnow.ExcelFormula, (ExcelFormulaContext)dateyearnow.ExcelFormula.Context, sheets);

                var dateyearnowResult = dateyearnowParser.Parse().LambdaInvoke<double>();
                System.Console.WriteLine($"dateyearnowResult = `{dateyearnowResult}`");

                var named = sheets[0].Rows[9].Cells[1];
                var namedParser = new ExpressionParser(named.ExcelFormula, (ExcelFormulaContext)named.ExcelFormula.Context, sheets);

                var namedResult = namedParser.Parse().LambdaInvoke<double>();
                System.Console.WriteLine($"namedResult = `{namedResult}`");


                var vlookup = sheets[3].Rows[0].Cells[4];
                var vlookupParser = new ExpressionParser(vlookup.ExcelFormula, (ExcelFormulaContext)vlookup.ExcelFormula.Context, sheets);
                var vlookupResult = vlookupParser.Parse().LambdaInvoke<double>();
                System.Console.WriteLine($"vlookupResult = `{vlookupResult}`");

                var vlookup2 = sheets[3].Rows[1].Cells[4];
                var vlookupParser2 = new ExpressionParser(vlookup2.ExcelFormula, (ExcelFormulaContext)vlookup2.ExcelFormula.Context, sheets);
                var vlookupResult2 = vlookupParser2.Parse().LambdaInvoke<double>();
                System.Console.WriteLine($"vlookupResult2 = `{vlookupResult2}`");

                int u = 0;
            }
        }

        private static XCell ToXCell(XRow row, ExcelRange range)
        {
            var cell = new XCell(row, range.Address)
            {
                FullAddress = $"{range.Worksheet.Name}!{range.Address}"
            };

            var context = new ExcelFormulaContext
            {
                Sheet = range.Worksheet.Name
            };

            if (!string.IsNullOrEmpty(range.Formula))
            {
                cell.Formula = "=" + range.Formula;
                cell.ExcelFormula = new ExcelFormula(cell.Formula, context);
            }
            else
            {
                cell.Value = range.Value;

                if (range.Value == null)
                {
                    return cell;
                }

                if (range.Value is bool)
                {
                    cell.ExcelFormula = new ExcelFormula((bool)range.Value ? "=TRUE" : "=FALSE", context);
                }
                else if (range.Value is DateTime)
                {
                    double value = DateTimeHelpers.ToOADate((DateTime) range.Value);
                    cell.ExcelFormula = new ExcelFormula("=" + value.ToString(CultureInfo.InvariantCulture), context);
                }
                else if (range.Value.IsNumeric())
                {
                    double value = Convert.ToDouble(range.Value);
                    cell.ExcelFormula = new ExcelFormula("=" + value.ToString(CultureInfo.InvariantCulture), context);
                }
                else
                {
                    cell.ExcelFormula = new ExcelFormula(string.Format("=\"{0}\"", range.Text), context);
                }
            }

            return cell;
        }

        private static void CalcTest()
        {
            // haakjes, machtsverheffen, vermenigvuldigen, delen, worteltrekken, optellen, aftrekken
            //var excelFormula = new ExcelFormula("=2^3 - -(1+1+1) * ROUND(4/2.7,2) + POWER(1+1,4) + 500 + SIN(3.1415926) + COS(3.1415926/2) + ABS(-1)");
            //var parser = new ExpressionParser(excelFormula);

            //Expression x = parser.Parse();
            //System.Console.WriteLine($"Expression = `{x}`");

            //var o = x.Optimize();

            //System.Console.WriteLine($"Expression = `{o}`");

            //LambdaExpression le = Expression.Lambda(o);

            //var result = le.Compile().DynamicInvoke();

            //System.Console.WriteLine($"result = `{result}`");

            //var trunc2 = new ExcelFormula("=TRUNC(91.789, 2)");
            //System.Console.WriteLine("trunc2 = `{0}`", Expression.Lambda(new ExpressionParser(trunc2).Parse()).Compile().DynamicInvoke());

            //var trunc0 = new ExcelFormula("=TRUNC(91.789)");
            //System.Console.WriteLine("trunc0 = `{0}`", Expression.Lambda(new ExpressionParser(trunc0).Parse()).Compile().DynamicInvoke());

            //var gt = new ExcelFormula("=1>2");
            //System.Console.WriteLine("gt = `{0}`", Expression.Lambda(new ExpressionParser(gt).Parse()).Compile().DynamicInvoke());

            //var sum = new ExcelFormula("=SUM(1,2,3)");
            //System.Console.WriteLine("sum = `{0}`", Expression.Lambda(new ExpressionParser(sum).Parse()).Compile().DynamicInvoke());

            var now = new ExcelFormula("=NOW()");
            System.Console.WriteLine("{0} : `{1}`", now.Formula, Expression.Lambda(new ExpressionParser(now).Parse()).Compile().DynamicInvoke());
        }
    }
}