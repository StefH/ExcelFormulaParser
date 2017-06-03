using System;
using System.Collections.Generic;

namespace ExcelFormulaParser.Expressions.Console
{
    public class XSheet
    {
        public string Name { get; private set; } 

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