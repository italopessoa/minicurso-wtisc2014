using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFactory
{
    public class RowStyle
    {
        #region Properties

        private IDictionary<int, NPOI.HSSF.UserModel.HSSFCellStyle> _cellStyles;
        private string _rowStyleName;
        private short? _rowForegroundColorIndex;
        private short? _rowFillPattern;

        public short? RowForegroundColor
        {
            get { return _rowForegroundColorIndex; }
            set { _rowForegroundColorIndex = value; }
        }

        public short? RowFillPattern
        {
            get { return _rowFillPattern; }
            set { _rowFillPattern = value; }
        }

        public IDictionary<int, NPOI.HSSF.UserModel.HSSFCellStyle> CellStyles
        {
            get { return _cellStyles; }
            private set { _cellStyles = value; }
        }

        public string RowStyleName
        {
            get { return _rowStyleName; }
            set { _rowStyleName = value; }
        }

        #endregion Properties

        public RowStyle()
        {
        }

        public void AddCellRowStyle(int cellIndex, NPOI.HSSF.UserModel.HSSFCellStyle cellStyle)
        {
            if (_cellStyles == null)
            {
                _cellStyles = new Dictionary<int, NPOI.HSSF.UserModel.HSSFCellStyle>();
            }

            if (_rowForegroundColorIndex.HasValue)
            {
                cellStyle.FillForegroundColor = _rowForegroundColorIndex.Value;
            }

            if (_rowFillPattern.HasValue)
            {
                cellStyle.FillPattern = _rowFillPattern.Value;
            }

            if (_cellStyles.Keys.Contains(cellIndex))
            {
                _cellStyles[cellIndex] = cellStyle;
            }
            else
            {
                _cellStyles.Add(cellIndex, cellStyle);
            }
        }
    }
}
