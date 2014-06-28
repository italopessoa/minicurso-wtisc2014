using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFactory
{
    public class ChildSheet : SpreadsheetFactory
    {
        private string _childPropertyName;

        public string ChildPropertyName
        {
            get { return _childPropertyName; }
        }
        
        public ChildSheet(string property)
        {
            _childPropertyName = property;
        }
    }
}
