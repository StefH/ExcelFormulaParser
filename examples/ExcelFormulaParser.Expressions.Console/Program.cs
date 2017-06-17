using System;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using ExcelFormulaExpressionParser;
using ExcelFormulaExpressionParser.Extensions;
using ExcelFormulaExpressionParser.Helpers;
using ExcelFormulaExpressionParser.Models;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Utilities;
using System.Reflection;
using log4net;
using log4net.Config;

namespace ExcelFormulaParser.Expressions.Console
{
    class Program
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof(Program));

        static void Main(string[] args)
        {
            var logRepository = LogManager.GetRepository(Assembly.GetEntryAssembly());
            XmlConfigurator.Configure(logRepository, new FileInfo("log4net.config"));

            Log.Info("Entering application.");
            Test();
            //CalcTest();
            //ExcelTest();
        }

        private static void Test()
        {
            using (var package = new ExcelPackage(new FileInfo("c:\\temp\\test.xlsx")))
            {
                var wb = new XWorkbook();

                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    var sheet = new XSheet(worksheet.Name);

                    // Obtain the worksheet size 
                    ExcelCellAddress endCell = worksheet.Dimension.End;

                    for (int r = 1; r <= endCell.Row; r++)
                    {
                        var xrow = new XRow(sheet, r);
                        for (int c = 1; c <= endCell.Column; c++)
                        {
                            xrow.Cells.Add(ToXCell(sheet, xrow, worksheet.Cells[r, c]));
                        }

                        sheet.Rows.Add(xrow);
                    }

                    wb.Sheets.Add(sheet);
                }

                var names = package.Workbook.Names;
                foreach (var name in names)
                {
                    wb.Names.Add(name.Name, name.Address);
                }

                var calcCell = wb.Sheets[2].Rows[16].Cells[2];
                var parser = new ExpressionParser(calcCell.ExcelFormula, 0, (ExcelFormulaContext)calcCell.ExcelFormula.Context, wb);

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
                var wb = new XWorkbook();

                package.Workbook.Worksheets[1].Cells["B1"].Value = 77;

                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    var sheet = new XSheet(worksheet.Name);

                    // Obtain the worksheet size 
                    ExcelCellAddress endCell = worksheet.Dimension.End;

                    for (int r = 1; r <= endCell.Row; r++)
                    {
                        var xrow = new XRow(sheet, r);
                        for (int c = 1; c <= endCell.Column; c++)
                        {
                            xrow.Cells.Add(ToXCell(sheet, xrow, worksheet.Cells[r, c]));
                        }

                        sheet.Rows.Add(xrow);
                    }

                    wb.Sheets.Add(sheet);
                }

                var names = package.Workbook.Names;
                foreach (var name in names)
                {
                    wb.Names.Add(name.Name, name.Address);
                }

                var calcCell = wb.Sheets[0].Rows[3].Cells[1];
                var parser = new ExpressionParser(calcCell.ExcelFormula, (ExcelFormulaContext)calcCell.ExcelFormula.Context, wb);

                Expression x = parser.Parse();
                System.Console.WriteLine($"Expression = `{x}`");

                var o = x.Optimize();

                System.Console.WriteLine($"Expression Optimize = `{o}`");

                LambdaExpression le = Expression.Lambda(o);

                var result = le.Compile().DynamicInvoke();

                System.Console.WriteLine($"result = `{result}`");

                var bool2 = wb.Sheets[0].Rows[4].Cells[1];
                var boolParser = new ExpressionParser(bool2.ExcelFormula, (ExcelFormulaContext)bool2.ExcelFormula.Context, wb);

                Expression bx = boolParser.Parse();
                System.Console.WriteLine($"Expression = `{bx}`");

                var o2 = bx.Optimize();

                System.Console.WriteLine($"Expression Optimize = `{o2}`");

                var bresult = Expression.Lambda(o2).Compile().DynamicInvoke();
                System.Console.WriteLine($"bresult = `{bresult}`");


                var sum = wb.Sheets[0].Rows[6].Cells[1];
                var sumParser = new ExpressionParser(sum.ExcelFormula, (ExcelFormulaContext)sum.ExcelFormula.Context, wb);

                Expression sumE = sumParser.Parse();
                System.Console.WriteLine($"Expression = `{sumE}`");

                var sumEOpt = sumE.Optimize();

                System.Console.WriteLine($"Expression Optimize = `{sumEOpt}`");

                var sumresult = Expression.Lambda(sumEOpt).Compile().DynamicInvoke();
                System.Console.WriteLine($"sumresult = `{sumresult}`");


                var sumSqrt = wb.Sheets[3].Rows[0].Cells[1];
                var sumSqrtParser = new ExpressionParser(sumSqrt.ExcelFormula, (ExcelFormulaContext)sumSqrt.ExcelFormula.Context, wb);

                Expression sumSqrtE = sumSqrtParser.Parse();
                //System.Console.WriteLine($"Expression = `{sumSqrtE}`");

                var sumSqrtEOpt = sumSqrtE.Optimize();

                //System.Console.WriteLine($"Expression Optimize = `{sumSqrtEOpt}`");

                var sumSqrtresult = Expression.Lambda(sumSqrtEOpt).Compile().DynamicInvoke();
                System.Console.WriteLine($"sumresult = `{sumSqrtresult}`");


                var dateyear = wb.Sheets[0].Rows[7].Cells[2];
                var dateyearParser = new ExpressionParser(dateyear.ExcelFormula, (ExcelFormulaContext)dateyear.ExcelFormula.Context, wb);

                Expression dateyearE = dateyearParser.Parse();
                //System.Console.WriteLine($"Expression = `{sumSqrtE}`");

                var dateyearO = dateyearE.Optimize();

                //System.Console.WriteLine($"Expression Optimize = `{sumSqrtEOpt}`");

                var dateyearResult = Expression.Lambda(dateyearO).Compile().DynamicInvoke();
                System.Console.WriteLine($"dateyearResult = `{dateyearResult}`");

                var dateyearnow = wb.Sheets[0].Rows[8].Cells[2];
                var dateyearnowParser = new ExpressionParser(dateyearnow.ExcelFormula, (ExcelFormulaContext)dateyearnow.ExcelFormula.Context, wb);

                var dateyearnowResult = dateyearnowParser.Parse().LambdaInvoke<double>();
                System.Console.WriteLine($"dateyearnowResult = `{dateyearnowResult}`");

                var named1 = wb.Sheets[0].Rows[9].Cells[1];
                var namedParser1 = new ExpressionParser(named1.ExcelFormula, (ExcelFormulaContext)named1.ExcelFormula.Context, wb);

                var namedResult1 = namedParser1.Parse().LambdaInvoke<double>();
                System.Console.WriteLine($"namedResult1 = `{namedResult1}`");

                var named2 = wb.Sheets[0].Rows[10].Cells[1];
                var namedParser2 = new ExpressionParser(named2.ExcelFormula, (ExcelFormulaContext)named2.ExcelFormula.Context, wb);

                var namedResult2 = namedParser2.Parse().LambdaInvoke<double>();
                System.Console.WriteLine($"namedResult2 = `{namedResult2}`");


                var vlookup = wb.Sheets[3].Rows[0].Cells[4];
                var vlookupParser = new ExpressionParser(vlookup.ExcelFormula, (ExcelFormulaContext)vlookup.ExcelFormula.Context, wb);
                var vlookupResult = vlookupParser.Parse().LambdaInvoke<double>();
                System.Console.WriteLine($"vlookupResult = `{vlookupResult}`");

                var vlookup2 = wb.Sheets[3].Rows[1].Cells[4];
                var vlookupParser2 = new ExpressionParser(vlookup2.ExcelFormula, (ExcelFormulaContext)vlookup2.ExcelFormula.Context, wb);
                var vlookupResult2 = vlookupParser2.Parse().LambdaInvoke<double>();
                System.Console.WriteLine($"vlookupResult2 = `{vlookupResult2}`");

                int u = 0;
            }
        }

        private static XCell ToXCell(XSheet sheet, XRow row, ExcelRange range)
        {
            var cell = new XCell(row, range.Address)
            {
                FullAddress = $"{range.Worksheet.Name}!{range.Address}"
            };

            var context = new ExcelFormulaContext
            {
                Sheet = sheet
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
                    cell.Expression = Expression.Constant((bool)range.Value);
                }
                else if (range.Value is DateTime)
                {
                    double value = DateTimeHelpers.ToOADate((DateTime)range.Value);
                    cell.Expression = Expression.Constant(value);
                }
                else if (range.Value.IsNumeric())
                {
                    double value = Convert.ToDouble(range.Value);
                    cell.Expression = Expression.Constant(value);
                }
                else
                {
                    cell.Expression = Expression.Constant(range.Text);
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