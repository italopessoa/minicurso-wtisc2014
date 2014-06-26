using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFactory
{
    public class Header
    {
        private int DEFAULT_TITLE_SPAN_SIZE = 10;

        private IDictionary<string, object> _filters;
        private string _title;
        private string _sheetName;

        public IDictionary<string, object> Filters
        {
            get { return _filters; }
        }

        public string Title
        {
            get { return _title; }
            set { _title = value; }
        }

        public string SheetName
        {
            get { return _sheetName; }
            set { _sheetName = value; }
        }

        private int TitleSpan
        {
            get { return DEFAULT_TITLE_SPAN_SIZE; }
            set { DEFAULT_TITLE_SPAN_SIZE = value; }
        }

        public bool AddFilter(string key, object value)
        {
            if (ValidadeHeaderValues(key, value))
            {
                if (_filters == null)
                {
                    _filters = new Dictionary<string, object>();
                }

                if (!_filters.Keys.Contains(key))
                {
                    _filters.Add(key, value);
                }

                return true;
            }
            return false;
        }

        private bool ValidadeHeaderValues(string key, object value)
        {
            return true;
        }

    }
}
