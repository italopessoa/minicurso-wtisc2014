using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFactory
{
    public class SpreadsheetFactory
    {
        private Header _header;

        private IList<TableHeader> _tableHeaders;

        private List<object> _datasource;
        
        private string[] _properties;

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

        public List<object> Datasource
        {
            get { return _datasource; }
            set { _datasource = value; }
        }

        public string[] Properties
        {
            get { return _properties; }
            set { _properties = value; }
        }
    }
}
