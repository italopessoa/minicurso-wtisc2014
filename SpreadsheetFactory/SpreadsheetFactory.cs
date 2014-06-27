using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFactory
{
    //TODO: definir um modelo para representar uma linha que realize o merge de acordo com a quantidade de celulas, servindo como titulo
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

        public int TableHeaderRows 
        {
            get { return SheetUtil.GetTotalHeaderRows(_tableHeaders); }
        }

        public int TableHeaderCells
        {
            get { return SheetUtil.GetTableHeaderCells(_tableHeaders); }
        }

        public int FirstHeaderCell { get; set; }

        private IList<SpreadsheetFactory> _spreadsheetFactoryList;

        public IList<SpreadsheetFactory> SpreadsheetFactoryList
        {
            get { return _spreadsheetFactoryList; }
            set { _spreadsheetFactoryList = value; }
        }

        public string Name { get; set; }

        public string MergedTitle { get; set; }
    }
}
