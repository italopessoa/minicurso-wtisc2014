using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFactory
{
    public class SpreadsheetFactory
    {
        private Header _header;

        private IList<TableHeader> _tableHeaders;

        public Header Header
        {
            get { return _header; }
            set { _header = value; }
        }

        public IList<TableHeader> TableHeaders
        {
            get { return _tableHeaders; }
            set { _tableHeaders = value; }
        }
    }
}
