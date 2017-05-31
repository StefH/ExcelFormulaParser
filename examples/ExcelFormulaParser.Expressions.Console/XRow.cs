using System.Collections.Generic;

namespace ExcelFormulaParser.Expressions.Console
{
    class XRow
    {
        public IList<XCell> Cells { get; set; }

        public XRow()
        {
            Cells = new List<XCell>();
        }
    }
}