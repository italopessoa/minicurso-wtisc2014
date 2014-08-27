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

        private IDictionary<string, List<ConditionalFormattingTemplate>> _conditionalFormatDictionary;

        public override IDictionary<string, List<ConditionalFormattingTemplate>> ConditionalFormatList
        {
            get { return _conditionalFormatDictionary; }
        }

        public override void AddConditionalFormatting(string property, ConditionalFormattingTemplate format)
        {
            if (_conditionalFormatDictionary == null)
            {
                _conditionalFormatDictionary = new Dictionary<string, List<ConditionalFormattingTemplate>>();
            }

            if (_conditionalFormatDictionary.Keys.Contains(property))
            {
                _conditionalFormatDictionary[property].Add(format);
                _conditionalFormatDictionary[property].Sort(delegate(ConditionalFormattingTemplate a, ConditionalFormattingTemplate b)
                {
                    return a.Priority.CompareTo(b.Priority);
                });
            }
            else
            {
                _conditionalFormatDictionary[property] = new List<ConditionalFormattingTemplate>();
                _conditionalFormatDictionary[property].Add(format);
            }
        }


    }
}
