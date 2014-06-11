using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFactory
{
    public class SpreadsheetFactory
    {
        private Header _header;
        
        private TableHeader _tableHeader;

        public Header Header
        {
            get { return _header; }
            set { _header = value; }
        }

        public TableHeader TableHeader
        {
            get { return _tableHeader; }
            set { _tableHeader = value; }
        }
    }
}
