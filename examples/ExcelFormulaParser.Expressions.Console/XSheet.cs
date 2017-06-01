using System.Collections.Generic;

namespace ExcelFormulaParser.Expressions.Console
{
    public class XSheet
    {
        public IList<XRow> Rows { get; set; }

        public XSheet()
        {
            Rows = new List<XRow>();
        }
    }
}