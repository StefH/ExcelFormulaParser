using System.Collections.Generic;
using System.Linq;
using ExcelFormulaExpressionParser.Models;

namespace ExcelFormulaExpressionParser.Utils
{
    internal class CellFinder
    {
        private readonly List<XSheet> _sheets;

        public CellFinder(List<XSheet> sheets)
        {
            _sheets = sheets;
        }

        public IEnumerable<XCell> Find(string sheetName, string address)
        {
            XSheet sheet;
            CellAddress start;
            CellAddress end;

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

                start = ExcelUtils.ParseExcelAddress(address);
                end = ExcelUtils.ParseExcelAddress(address);

                //return new[] {sheet.Rows.SelectMany(r => r.Cells).FirstOrDefault(c => c.Address == address)};
            }
            else
            {
                string[] partsRange = address.Split(':');


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
            }

            var rows = sheet.Rows.GetRange(start.Row - 1, end.Row - start.Row + 1);

            return rows.SelectMany(r => r.Cells)
               .Where(c => c.Column >= start.Column && c.Column <= end.Column);
        }
    }
}