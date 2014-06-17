using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFactory
{
    public class ConditionalFormattingTemplate
    {
        public string Message { get; set; }
        public string PropertyName { get; set; }
        public object Value { get; set; }
        public NPOI.HSSF.Record.ComparisonOperator ComparisonOperator { get; set; }
        public NPOI.HSSF.UserModel.HSSFCellStyle CellStyle { get; set; }
    }
}
