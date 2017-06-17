using System;
using System.Linq;
using ExcelFormulaExpressionParser.Models;

namespace ExcelFormulaExpressionParser.Utils
{
    internal class CellFinder : ICellFinder
    {
        private readonly XWorkbook _workbook;

        private readonly bool _hasNames;

        public CellFinder(XWorkbook workbook)
        {
            _workbook = workbook;
            _hasNames = workbook.Names.Count > 0;
        }

        public XRange Find(XSheet sheet, string address)
        {
            if (_hasNames && _workbook.Names.ContainsKey(address))
            {
                address = _workbook.Names[address];
            }

            address = address.Replace("$", "");

            var range = new XRange();

            if (!address.Contains(':'))
            {
                if (!address.Contains('!'))
                {
                    // Same sheet
                    range.Sheet = _workbook.Sheets.First(s => s.Name == sheet.Name);
                }
                else
                {
                    // Other sheet
                    string[] parts = address.Split('!');
                    range.Sheet = _workbook.Sheets.First(s => s.Name == parts[0]);
                    address = parts[1];
                }

                range.Start = ExcelUtils.ParseExcelAddress(address);
                range.End = ExcelUtils.ParseExcelAddress(address);
            }
            else
            {
                string[] partsRange;
                if (!address.Contains('!'))
                {
                    // Same sheet
                    partsRange = address.Split(':');
                    range.Sheet = _workbook.Sheets.First(s => s.Name == sheet.Name);
                }
                else
                {
                    // Other sheet
                    string[] parts = address.Split('!');
                    range.Sheet = _workbook.Sheets.First(s => s.Name == parts[0]);
                    address = parts[1];

                    partsRange = address.Split(':');
                }

                range.Start = ExcelUtils.ParseExcelAddress(partsRange[0]);
                range.End = ExcelUtils.ParseExcelAddress(partsRange[1]);
            }

            try
            {
                var rows = range.Sheet.Rows.GetRange(range.Start.Row - 1, range.End.Row - range.Start.Row + 1);

                range.Cells = rows.SelectMany(r => r.Cells)
                    .Where(c => c.Column >= range.Start.Column && c.Column <= range.End.Column).ToArray();
            }
            catch
            {
                range.Cells = new[] { new XCell(new XRow(range.Sheet, range.Start.Row), address) };
            }

            range.Address = address;
            return range;
        }
    }
}