using System.Collections.Generic;
using System.Linq;
using ExcelFormulaExpressionParser.Models;

namespace ExcelFormulaExpressionParser.Utils
{
    internal class CellFinder : ICellFinder
    {
        private readonly List<XSheet> _sheets;

        public CellFinder(List<XSheet> sheets)
        {
            _sheets = sheets;
        }

        public XRange Find(string sheetName, string address)
        {
            address = address.Replace("$", "");

            var range = new XRange();

            if (!address.Contains(':'))
            {
                if (!address.Contains('!'))
                {
                    // Same sheet
                    range.Sheet = _sheets.First(s => s.Name == sheetName);
                }
                else
                {
                    // Other sheet
                    string[] parts = address.Split('!');
                    range.Sheet = _sheets.First(s => s.Name == parts[0]);
                    address = parts[1];
                }

                range.Start = ExcelUtils.ParseExcelAddress(address);
                range.End = ExcelUtils.ParseExcelAddress(address);
            }
            else
            {
                string[] partsRange = address.Split(':');

                if (!address.Contains('!'))
                {
                    // Same sheet
                    range.Sheet = _sheets.First(s => s.Name == sheetName);
                    range.Start = ExcelUtils.ParseExcelAddress(partsRange[0]);
                    range.End = ExcelUtils.ParseExcelAddress(partsRange[1]);
                }
                else
                {
                    // Other sheet
                    string[] partsStart = partsRange[0].Split('!');
                    string[] partEnd = partsRange[1].Split('!');

                    range.Sheet = _sheets.First(s => s.Name == partsStart[0]);
                    range.Start = ExcelUtils.ParseExcelAddress(partsStart[1]);
                    range.End = ExcelUtils.ParseExcelAddress(partEnd[1]);
                }
            }

            var rows = range.Sheet.Rows.GetRange(range.Start.Row - 1, range.End.Row - range.Start.Row + 1);

            range.Cells = rows.SelectMany(r => r.Cells)
               .Where(c => c.Column >= range.Start.Column && c.Column <= range.End.Column).ToArray();

            range.Address = address;
            return range;
        }
    }
}