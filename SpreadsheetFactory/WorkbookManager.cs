using NPOI.HSSF.Record;
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
        protected static int tableHeaderRows = 0;

        private static int _tableHeaderCells = 0;

        private static int headerCell;

        private static HSSFWorkbook _workbook;

        private static HSSFCellStyle _headerCellStyle = null;

        private static HSSFCellStyle _defaultContentCellStyle = null;

        private static IDictionary<string, List<ConditionalFormattingTemplate>> _conditionalFormatDictionary;

        static WorkbookManager()
        {
            if (_workbook == null)
            {
                _workbook = new HSSFWorkbook();
            }
        }

        public static void AddConditionalFormatting(string property, ConditionalFormattingTemplate format)
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
                //if(_conditionalFormatDictionary[property] == null)
                //{
                _conditionalFormatDictionary[property] = new List<ConditionalFormattingTemplate>();
                _conditionalFormatDictionary[property].Add(format);
                //}
            }
        }

        public static HSSFCellStyle GetNewHSSFCellStyle()
        {
            return _workbook.CreateCellStyle();
        }

        public static HSSFFont GetNewHSSFCellFont()
        {
            return _workbook.CreateFont();
        }

        public static HSSFDataFormat GetNewHSSFDataFormat()
        {
            return _workbook.CreateDataFormat();
        }
        public static void SetHeaderCellStyle(HSSFCellStyle cellStyle)
        {
            _headerCellStyle = cellStyle;
        }

        public static void SetDefaultContentCellStyle(HSSFCellStyle cellStyle)
        {
            _defaultContentCellStyle = cellStyle;
        }

        public static void CreateSpreadsheet(SpreadsheetFactory spreadsheetFactory)
        {
            if (spreadsheetFactory != null)
            {
                HSSFSheet sheet = null;
                HSSFPatriarch drawingPatriarch = null;
                if (spreadsheetFactory.Header != null)
                {
                    sheet = CreateTitle(spreadsheetFactory.Header.Title, spreadsheetFactory.Header.SheetName);
                    drawingPatriarch = sheet.CreateDrawingPatriarch();

                    if (spreadsheetFactory.Header.Filters != null && spreadsheetFactory.Header.Filters.Count > 0)
                    {
                        sheet = ConfigHeader(sheet, spreadsheetFactory.Header.Filters);
                    }

                    sheet.CreateRow(sheet.LastRowNum + 1);
                    sheet.CreateRow(sheet.LastRowNum + 1);

                    PrepareTableHeader(sheet, spreadsheetFactory.TableHeaders, 0);
                    int cellAux = 0;
                    int firstRow = (sheet.LastRowNum - tableHeaderRows);
                    ConfigTableHeader(sheet, spreadsheetFactory.TableHeaders, ref cellAux, (sheet.LastRowNum - tableHeaderRows));

                    if (_headerCellStyle != null)
                    {
                        ApplyHeaderCellStyle(sheet, firstRow, cellAux, tableHeaderRows, _tableHeaderCells, _headerCellStyle);
                    }

                }

                if (spreadsheetFactory.Datasource != null)
                {
                    if (sheet == null)
                    {
                        sheet = _workbook.CreateSheet();
                    }

                    int firstRow = sheet.LastRowNum;
                    CreateContentTableCells(sheet, spreadsheetFactory.Datasource.Count, spreadsheetFactory.Datasource.Count, 0, firstRow);
                    SetContentTableCellsValue(sheet, firstRow, 0, spreadsheetFactory.Properties, spreadsheetFactory.Datasource);

                    if (_conditionalFormatDictionary != null)
                    {
                        List<PropertyCell> teste = new List<PropertyCell>();
                        for (int i = 0; i < spreadsheetFactory.Properties.Length; i++)
                        {
                            if (_conditionalFormatDictionary.Keys.Contains(spreadsheetFactory.Properties[i]))
                            {
                                teste.Add(new PropertyCell() { CellIndex = i, PropertyName = spreadsheetFactory.Properties[i] });
                            }
                        }

                        //ApplyContentListCellStyle(sheet, firstRow, 0, spreadsheetFactory.Datasource.Count, spreadsheetFactory.Properties.Length, _defaultContentCellStyle, drawingPatriarch);
                        ApplyConditinalFormattingContentTable(sheet, firstRow, spreadsheetFactory.Datasource.Count, 0, spreadsheetFactory.Properties.Length + 0, teste);
                    }

                    #region configurar formatacao condicional
                    //HSSFSheetConditionalFormatting cf = sheet.SheetConditionalFormatting;
                    //HSSFConditionalFormattingRule rule = cf.CreateConditionalFormattingRule(NPOI.HSSF.Record.ComparisonOperator.GE, "3", null);

                    //HSSFPatternFormatting patternFmt = rule.CreatePatternFormatting();
                    //patternFmt.FillBackgroundColor = HSSFColor.RED.index;
                    //CellRangeAddress[] range =
                    //{
                    //    new CellRangeAddress(9,12,0,0)
                    //};
                    //cf.AddConditionalFormatting(range, rule);
                    #endregion configurar formatacao condicional

                }
            }
        }

        private static void ApplyHeaderCellStyle(HSSFSheet sheet, int firstRow, int firstCell, int rows, int cells, HSSFCellStyle style)
        {
            for (int r = firstRow; r < (rows + firstRow); r++)
            {
                for (int c = firstCell; c < (cells + firstCell); c++)
                {
                    sheet.GetRow(r).GetCell(c).CellStyle = style;
                }
            }
        }

        private static void ApplyContentListCellStyle(HSSFSheet sheet, int firstRow, int firstCell, int rows, int cells, HSSFCellStyle style, HSSFPatriarch drawingPatriarch)
        {
            HSSFCellStyle templateStyle23 = WorkbookManager.GetNewHSSFCellStyle();
            templateStyle23.BorderTop = 2;
            templateStyle23.BorderRight = 2;
            templateStyle23.BorderRight = 2;
            templateStyle23.BorderBottom = 2;
            templateStyle23.FillPattern = HSSFCellStyle.SOLID_FOREGROUND;
            templateStyle23.FillForegroundColor = HSSFColor.PINK.index;

            for (int r = firstRow; r < (rows + firstRow); r++)
            {
                for (int c = firstCell; c < (cells + firstCell); c++)
                {
                    sheet.GetRow(r).RowStyle = templateStyle23;
                }
            }
        }

        private static void ApplyConditinalFormattingContentTable(HSSFSheet sheet, int firstRow, int rows, int firstCell, int lastCell, List<PropertyCell> cellPoints)
        {
            for (int r = firstRow; r < (firstRow + rows); r++)
            {
                RowStyle style = TesteA(sheet.GetRow(r), cellPoints);

                if (style != null)
                {
                    ApplyRowStyle(sheet.GetRow(r), style);
                }

                //for (int c = firstCell; c < lastCell; c++)
                //{
                //    .GetCell(c).CellStyle = style;
                //}
            }
        }

        private static RowStyle TesteA(HSSFRow row, List<PropertyCell> cellPoints)
        {
            ConditionalFormattingTemplate template = null;
            foreach (var item in cellPoints)
            {
                foreach (var style in _conditionalFormatDictionary[item.PropertyName])
                {
                    if (IsMatchStyle(style, row.GetCell(item.CellIndex)))
                    {
                        if (template != null)
                        {
                            if (style.Priority > template.Priority)
                            {
                                //System.Diagnostics.Debug.WriteLine("A: "+template.RowStyle.RowStyleName + ">" + template.RowStyle.RowForegroundColor);
                                template = style;
                                //System.Diagnostics.Debug.WriteLine("D: "+template.RowStyle.RowStyleName + ">" + template.RowStyle.RowForegroundColor);
                            }
                        }
                        else
                        {
                            template = style;
                        }
                    }
                }
            }
            if (template != null)
            {
                //System.Diagnostics.Debug.WriteLine(template.RowStyle.RowStyleName + ">" + template.RowStyle.RowForegroundColor);
                return template.RowStyle;
            }
            else
            {
                return null;
            }
        }

        private static bool IsMatchStyle(ConditionalFormattingTemplate cft, HSSFCell cell)
        {
            int CELL_TYPE_INDEX = SheetUtil.GetCellType(cft.Value);

            switch (cell.CellType)
            {
                case NPOI.HSSF.UserModel.HSSFCell.CELL_TYPE_NUMERIC:
                    break;
                case NPOI.HSSF.UserModel.HSSFCell.CELL_TYPE_STRING:
                    break;
                case SheetUtil.CELL_TYPE_DATETIME:
                    break;
            }

            // >, <, ==, empty, null
            switch (cft.ComparisonOperator)
            {
                case ComparisonOperator.BETWEEN:
                    break;
                case ComparisonOperator.EQUAL: //igual
                    switch (CELL_TYPE_INDEX)
                    {
                        case NPOI.HSSF.UserModel.HSSFCell.CELL_TYPE_NUMERIC:
                            double na = cell.NumericCellValue;
                            double nb;
                            double.TryParse(cft.Value.ToString(), out nb);
                            return na == nb;
                        case NPOI.HSSF.UserModel.HSSFCell.CELL_TYPE_STRING:
                            string sa = cell.StringCellValue;
                            string sb = (string)cft.Value;
                            return sa.Equals(sb);
                        case SheetUtil.CELL_TYPE_DATETIME:
                            DateTime da = cell.DateCellValue;
                            DateTime db = (DateTime)cft.Value;
                            return da == db;
                    }
                    break;
                case ComparisonOperator.GE:
                    break;
                case ComparisonOperator.GT: //maior
                    switch (CELL_TYPE_INDEX)
                    {
                        case NPOI.HSSF.UserModel.HSSFCell.CELL_TYPE_NUMERIC:
                            double na = cell.NumericCellValue;
                            double nb;
                            double.TryParse(cft.Value.ToString(), out nb);
                            return na > nb;
                        case SheetUtil.CELL_TYPE_DATETIME:
                            DateTime da = cell.DateCellValue;
                            DateTime db = (DateTime)cft.Value;
                            return da > db;
                    }
                    break;
                case ComparisonOperator.LE:
                    break;
                case ComparisonOperator.LT: //menor
                    switch (CELL_TYPE_INDEX)
                    {
                        case NPOI.HSSF.UserModel.HSSFCell.CELL_TYPE_NUMERIC:
                            double na = cell.NumericCellValue;
                            double nb;
                            double.TryParse(cft.Value.ToString(), out nb);
                            return na < nb;
                        case SheetUtil.CELL_TYPE_DATETIME:
                            DateTime da = cell.DateCellValue;
                            DateTime db = (DateTime)cft.Value;
                            return da < db;
                    }
                    break;
                case ComparisonOperator.NOT_BETWEEN:
                    break;
                case ComparisonOperator.NOT_EQUAL:
                    break;
                case ComparisonOperator.NO_COMPARISON:
                    break;
                default:
                    break;
            }

            return false;

        }

        //private static HSSFCellStyle GetConditionalStyle(List<PropertyCell> cellPoints)
        //{
        //    int maxPriority = 0;
        //    int cellToComment = 0;

        //    foreach (var item in cellPoints)
        //    {
        //        foreach (var s in _conditionalFormatDictionary[item.PropertyName])
        //        {
        //            s.Priority = (s.Priority > maxPriority ? s.Priority : maxPriority);
        //        }
        //    }
        //}

        private static void AddConditionalFormatComment(HSSFSheet sheet, HSSFCell cell, string comment)
        {
            /*HSSFPatriarch patr = sheet.DrawingPatriarch CreateDrawingPatriarch();
            //anchor defines size and position of the comment in worksheet
            //row1 linha final do comentario, 1 pe como um valor padrao para a altura do comentario
            //col2 é a quantidade de colunas que o comentario irá abranger
            //row2 é a quantidade de linhas que o comentario tem
            HSSFComment comment1 = patr.CreateComment(new HSSFClientAnchor(100, 100, 100, 100, (short)1, 1, (short)6, 5));
            comment1.String = new HSSFRichTextString("FirstComments");
            sheet.GetRow(r).GetCell(c).CellComment = comment1;*/
        }

        private static void PrepareTableHeader(HSSFSheet sheet, IList<TableHeader> tableHeaders, int cell)
        {
            tableHeaderRows = SheetUtil.GetTotalHeaderRows(tableHeaders);
            CreateHeaderRows(sheet, tableHeaders);
            GetTableHeaderCells(tableHeaders);
            CreateHeaderCells(sheet, cell);
        }

        private static void ConfigTableHeader(HSSFSheet sheet, IList<TableHeader> tableHeaders, ref int cell, int row)
        {
            headerCell = cell;
            int cellRow = row;

            foreach (var item in tableHeaders)
            {
                if (item.Cells == null)
                {
                    sheet.GetRow(cellRow).GetCell(headerCell).SetCellValue(item.Text);

                    //mesclar linhas
                    sheet.AddMergedRegion(new CellRangeAddress(cellRow, sheet.LastRowNum - 1, headerCell, headerCell + item.SpanSize));
                    headerCell++;
                }
                else if (item.Cells != null)
                {
                    sheet.GetRow(cellRow).GetCell(headerCell).SetCellValue(item.Text);

                    int headerCellAux = headerCell + item.Cells.Count - 1;

                    //mesclar celulas
                    sheet.AddMergedRegion(new CellRangeAddress(cellRow, cellRow, headerCell, headerCell + item.SpanSize));

                    bool callRecursive = true;
                    foreach (var internItem in item.Cells)
                    {
                        if (internItem.Cells == null || internItem.Cells.Count == 0)
                        {
                            sheet.GetRow(cellRow + 1).GetCell(headerCell).SetCellValue(internItem.Text);
                            //mesclar linhas
                            int rowAux = cellRow + 1;
                            if (rowAux > (sheet.LastRowNum - 1))
                            {
                                rowAux = sheet.LastRowNum - 1;
                            }
                            sheet.AddMergedRegion(new CellRangeAddress(cellRow + 1, sheet.LastRowNum - 1, headerCell, headerCell));
                            headerCell++;
                            callRecursive = false;
                        }
                    }

                    if (callRecursive)
                    {
                        ConfigTableHeader(sheet, item.Cells, ref headerCell, cellRow + 1);
                    }
                }
            }
        }

        private static void CreateHeaderRows(HSSFSheet sheet, IList<TableHeader> tableHeaders)
        {
            for (int i = 0; i < tableHeaderRows; i++)
            {
                sheet.CreateRow(sheet.LastRowNum + 1);
            }
        }

        private static void CreateHeaderCells(HSSFSheet sheet, int firstCell)
        {
            for (int i = 1; i < tableHeaderRows + 1; i++)
            {
                for (int j = firstCell; j < _tableHeaderCells + firstCell; j++)
                {
                    //sheet.GetRow(sheet.LastRowNum - i).CreateCell(j).SetCellValue(i + ":" + j);
                    sheet.GetRow(sheet.LastRowNum - i).CreateCell(j);
                }
                Console.WriteLine(sheet.LastRowNum - i);
            }
        }

        private static void CreateContentTableCells(HSSFSheet sheet, int rows, int cells, int firstCell, int firstRow)
        {
            for (int i = sheet.LastRowNum; i < firstRow + rows; i++)
            {
                sheet.CreateRow(i);
                for (int j = firstCell; j < firstCell + cells; j++)
                {
                    sheet.GetRow(i).CreateCell(j);//.SetCellValue("A");
                }
            }
        }

        private static void SetContentTableCellsValue(HSSFSheet sheet, int firstRow, int firstCell, string[] properties, List<object> values)
        {
            int row = firstRow;
            object propValue;

            foreach (var item in values)
            {
                for (int i = firstCell; i < properties.Length; i++)
                {
                    propValue = GetPropValue(item, properties[i]);
                    switch (SheetUtil.GetCellType(propValue))
                    {
                        case HSSFCell. CELL_TYPE_NUMERIC:
                            sheet.GetRow(row).GetCell(i).SetCellValue( double.Parse(propValue.ToString()));
                            break;
                        case HSSFCell.CELL_TYPE_STRING:
                            sheet.GetRow(row).GetCell(i).SetCellValue(propValue.ToString());
                            break;
                        case SheetUtil.CELL_TYPE_DATETIME:
                            sheet.GetRow(row).GetCell(i).SetCellValue(((DateTime)propValue));
                            break;
                    }
                }
                row++;
            }
        }

        private static object GetPropValue(object src, string propName)
        {
            return src.GetType().GetProperty(propName).GetValue(src, null);
        }

        private static void GetTableHeaderCells(IList<TableHeader> tableHeaders)
        {
            foreach (var item in tableHeaders)
            {
                if (item.Cells == null || item.Cells.Count == 0)
                {
                    _tableHeaderCells++;
                }
                else
                {
                    //tableHeaderCells++;
                    GetTableHeaderCells(item.Cells);
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

                switch (SheetUtil.GetCellType(item.Value))
                {
                    case HSSFCell.CELL_TYPE_NUMERIC:
                        sheet.GetRow(sheet.LastRowNum).CreateCell(1, HSSFCell.CELL_TYPE_NUMERIC).SetCellValue(double.Parse(item.Value.ToString()));
                        break;
                    case HSSFCell.CELL_TYPE_STRING:
                        sheet.GetRow(sheet.LastRowNum).CreateCell(1, HSSFCell.CELL_TYPE_STRING).SetCellValue(item.Value.ToString());
                        break;
                    case SheetUtil.CELL_TYPE_DATETIME:
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

            _workbook = null;
        }

        private static void ApplyRowStyle(HSSFRow row, RowStyle rowStyle)
        {
            foreach (KeyValuePair<int, HSSFCellStyle> style in rowStyle.CellStyles)
            {
                System.Diagnostics.Debug.WriteLine("XXXXXXXXXX >>>> "+rowStyle.RowStyleName + ">" + rowStyle.RowForegroundColor);
                row.GetCell(style.Key + row.FirstCellNum).CellStyle = style.Value;
            }
        }

        internal struct PropertyCell
        {
            public int CellIndex { get; set; }
            public string PropertyName { get; set; }
        }
    }

    public class Pessoa
    {
        public string Nome { get; set; }
        public int Idade { get; set; }
        public DateTime Nascimento { get; set; }
        public Double Salario { get; set; }
    }
}
