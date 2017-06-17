using System.Collections.Generic;

namespace ExcelFormulaExpressionParser.Models
{
    public class XWorkbook
    {
        public List<XSheet> Sheets { get; set; }

        public IDictionary<string, string> Names { get; set; }

        public XWorkbook()
        {
            Names = new Dictionary<string, string>();
            Sheets = new List<XSheet>();
        }
    }
}