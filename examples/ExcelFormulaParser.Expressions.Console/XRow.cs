using System.Collections.Generic;
using System.Linq;

namespace ExcelFormulaParser.Expressions.Console
{
    public class XRow
    {
        private XSheet Sheet { get; }
        public IList<XCell> Cells { get; set; }

        public XRow(XSheet sheet)
        {
            Sheet = sheet;
            Cells = new List<XCell>();
        }

        //public XCell this[string address]
        //{
        //    get
        //    {
        //        if (!address.Contains('!'))
        //        {
        //            return Cells.FirstOrDefault(c => c.Address == address);
        //        }
                
        //        {
        //            var rows = sheets.SelectMany(s => s.Rows);
        //            return rows.SelectMany(r => r.Cells).FirstOrDefault(c => c.FullAddress == address);
        //        }
        //    }
        //}
    }
}