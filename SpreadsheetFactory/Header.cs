using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFactory
{
    public class Header
    {
        private IDictionary<string, object> _filters;
        private string _title;

        public IDictionary<string,object> Filters
        {
            get { return _filters; }
            //set { _filters = value; }
        }

        public string Title
        {
            get { return _title; }
            set { _title = value; }
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
