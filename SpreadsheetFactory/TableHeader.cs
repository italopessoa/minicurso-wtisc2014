using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFactory
{
    public class TableHeader : WorkbookManager
    {
        private string _text;
        private IList<TableHeader> _cells;

        public string Text
        {
            get { return _text; }
            set { _text = value; }
        }

        public IList<TableHeader> Cells
        {
            get { return _cells; }
            set { _cells = value; }
        }

        public void AddSpanCell(string value)
        {
            _cells.Add(new TableHeader() { _text = value });
        }
    }
}
