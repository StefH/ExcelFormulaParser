using System;
using System.Collections.Generic;

namespace ExcelFormulaExpressionParser
{
    public class XSheet
    {
        public string Name { get; } 

        public IList<XRow> Rows { get; set; }

        public XSheet(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentNullException(nameof(name));
            }

            Name = name;
            Rows = new List<XRow>();
        }
    }
}