using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFactory
{
    public class TableHeader// : WorkbookManager
    {
        private string _text;
        private IList<TableHeader> _cells = null;

        public string Text
        {
            get { return _text; }
            set
            {
                _text = value;
            }
        }

        public IList<TableHeader> Cells
        {
            get { return _cells; }
            set { _cells = value; }
        }

        public TableHeader AddSpanCell(string value)
        {
            TableHeader tableHeader = new TableHeader();
            tableHeader.Text = value;
            if (_cells == null)
            {
                _cells = new List<TableHeader>();
            }
            _cells.Add(tableHeader);


            return tableHeader;
        }

        private int CellSpanCount(IList<TableHeader> list)
        {
            int total = 0;
            foreach (var item in list)
            {
                if (item.Cells == null)
                {
                    total++;
                }
                else
                {
                    total += CellSpanCount(item.Cells);
                }
            }

            return total;
        }

        public int SpanSize
        {
            get
            {
                if (_cells != null)
                {
                    return CellSpanCount(_cells) - 1;
                }
                else
                {
                    return 0;
                }
            }
        }
    }
}
