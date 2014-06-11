using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SpreadsheetFactory
{
    public class WorkbookManager
    {
        private static HSSFWorkbook _workbook;

        private static HSSFWorkbook Workbook
        {
            get
            {
                if (_workbook == null)
                {
                    _workbook = new HSSFWorkbook();
                }
                return _workbook;
            }
        }

        static WorkbookManager()
        {
            if (_workbook == null)
            {
                _workbook = new HSSFWorkbook();
            }
        }

        public static void CreateSpreadsheet(SpreadsheetFactory spreadsheetFactory)
        {
            if (spreadsheetFactory != null)
            {
                if (spreadsheetFactory.Header != null)
                {
                    HSSFSheet sheet = CreateTitle(spreadsheetFactory.Header.Title, spreadsheetFactory.Header.SheetName);

                    if (spreadsheetFactory.Header.Filters != null && spreadsheetFactory.Header.Filters.Count > 0)
                    {
                        sheet = ConfigHeader(sheet, spreadsheetFactory.Header.Filters);
                    }
                }
            }
        }

        private static HSSFSheet CreateTitle(string title, string sheetName)
        {
            HSSFSheet newSheet;

            if (!string.IsNullOrEmpty(sheetName))
            {
                newSheet = _workbook.CreateSheet(sheetName);
            }
            else
            {
                newSheet = _workbook.CreateSheet();
            }

            if (!string.IsNullOrEmpty(title))
            {
                newSheet.CreateRow(newSheet.LastRowNum).CreateCell(0).SetCellValue(title);
                newSheet.AddMergedRegion(new CellRangeAddress(newSheet.LastRowNum, newSheet.LastRowNum, 0, SheetUtil.DEFAULT_TITLE_SPAN_SIZE));
            }

            return newSheet;
        }

        private static HSSFSheet ConfigHeader(HSSFSheet sheet, IDictionary<string, object> filters)
        {
            foreach (KeyValuePair<string, object> item in filters)
            {
                sheet.CreateRow(sheet.LastRowNum + 1).CreateCell(0).SetCellValue(item.Key);

                Type type = item.Value.GetType();

                switch (type.FullName)
                {
                    case "System.Int16":
                    case "System.Int32":
                    case "System.Int64":
                    case "System.Double":
                        sheet.GetRow(sheet.LastRowNum).CreateCell(1, HSSFCell.CELL_TYPE_NUMERIC).SetCellValue(double.Parse(item.Value.ToString()));
                        break;
                    case "System.String":
                        sheet.GetRow(sheet.LastRowNum).CreateCell(1, HSSFCell.CELL_TYPE_STRING).SetCellValue(item.Value.ToString());
                        break;
                    case "System.DateTime":
                        sheet.GetRow(sheet.LastRowNum).CreateCell(1, HSSFCell.CELL_TYPE_STRING).SetCellValue(((DateTime)item.Value).ToShortDateString());
                        break;
                }

            }

            return sheet;
        }

        public static void SaveSpreadsheet(string filePath, string fileName)
        {
            MemoryStream ms = new MemoryStream();
            _workbook.Write(ms);
            FileStream fs = new FileStream(filePath + fileName, FileMode.Create);
            fs.Write(ms.ToArray(), 0, Convert.ToInt32(ms.Length));
            fs.Close();
            ms.Close();
        }
    }
}
