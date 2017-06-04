using System.Collections.Generic;
using System.Linq;
using ExcelFormulaExpressionParser.Models;

namespace ExcelFormulaExpressionParser.Utils
{
    internal class CellFinder
    {
        private readonly IList<XSheet> _sheets;

        public CellFinder(IList<XSheet> sheets)
        {
            _sheets = sheets;
        }

        public XCell[] Find(string sheetName, string address)
        {
            XSheet sheet;
            if (!address.Contains(':'))
            {
                if (!address.Contains('!'))
                {
                    // Same sheet
                    sheet = _sheets.First(s => s.Name == sheetName);
                }
                else
                {
                    // Other sheet
                    string[] parts = address.Split('!');
                    sheet = _sheets.First(s => s.Name == parts[0]);
                    address = parts[1];
                }

                return new[] { sheet.Rows.SelectMany(r => r.Cells).FirstOrDefault(c => c.Address == address) };
            }

            string[] partsRange = address.Split(':');
            CellAddress start;
            CellAddress end;

            if (!address.Contains('!'))
            {
                // Same sheet
                sheet = _sheets.First(s => s.Name == sheetName);
                start = ExcelUtils.ParseExcelAddress(partsRange[0]);
                end = ExcelUtils.ParseExcelAddress(partsRange[1]);
            }
            else
            {
                // Other sheet
                string[] partsStart = partsRange[0].Split('!');
                string[] partEnd = partsRange[1].Split('!');

                sheet = _sheets.First(s => s.Name == partsStart[0]);
                start = ExcelUtils.ParseExcelAddress(partsStart[1]);
                end = ExcelUtils.ParseExcelAddress(partEnd[1]);
            }

            return sheet.Rows
                .SelectMany(r => r.Cells)
                .Where(c =>
                    c.Column >= start.Column && c.Column <= end.Column &&
                    c.Row >= start.Row && c.Row <= end.Row
                )
                .ToArray();
        }
    }
}