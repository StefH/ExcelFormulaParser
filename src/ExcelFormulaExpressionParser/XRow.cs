using System.Collections.Generic;

namespace ExcelFormulaExpressionParser
{
    public class XRow
    {
        private XSheet Sheet { get; }

        public int Row { get; }

        public List<XCell> Cells { get; set; }

        public XRow(XSheet sheet, int row)
        {
            Sheet = sheet;
            Row = row;
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