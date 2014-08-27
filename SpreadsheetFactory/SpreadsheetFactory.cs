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

        private IList<object> _datasource;
        
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

        public IList<object> Datasource
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

        public NPOI.HSSF.UserModel.HSSFCellStyle HeaderCellStyle { get; set; }

        public RowStyle RowStyle { get; set; }

        private IDictionary<string, List<ConditionalFormattingTemplate>> _conditionalFormatDictionary;

        public virtual IDictionary<string, List<ConditionalFormattingTemplate>> ConditionalFormatList
        {
            get { return _conditionalFormatDictionary; }
        }

        public virtual void AddConditionalFormatting(string property, ConditionalFormattingTemplate format)
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

        //private SpreadsheetFactory _childSheet;
        private ChildSheet _childSheet;

        public ChildSheet ChildSheet
        {
            get { return _childSheet; }
            set { _childSheet = value; }
        }

        private int _firstCell = 0;

        public int FirstCell
        {
            get { return _firstCell; }
            set { _firstCell =  value; }
        }

        private NPOI.HSSF.UserModel.HSSFCellStyle _spanTitleStyle;

        public NPOI.HSSF.UserModel.HSSFCellStyle SpanTitleStyle
        {
            get { return _spanTitleStyle; }
            set {  _spanTitleStyle = value; }
        }
    }
}
