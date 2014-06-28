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
        #region Private Members

        private int headerCell;

        private HSSFWorkbook _workbook;

        #endregion Private Members

        #region Public Methods

        public HSSFCellStyle GetNewHSSFCellStyle()
        {
            return _workbook.CreateCellStyle();
        }

        public HSSFFont GetNewHSSFCellFont()
        {
            return _workbook.CreateFont();
        }

        public HSSFDataFormat GetNewHSSFDataFormat()
        {
            return _workbook.CreateDataFormat();
        }

        public void MountSpreadsheet(SpreadsheetFactory spreadsheetFactory)
        {
            if (spreadsheetFactory.SpreadsheetFactoryList != null)
            {
                foreach (var item in spreadsheetFactory.SpreadsheetFactoryList)
                {
                    CreateSpreadsheet(item);
                }
            }
        }

        private void CreateSpreadsheet(SpreadsheetFactory spreadsheetFactory)
        {
            if (spreadsheetFactory != null)
            {
                //TODO: verificar se é possivel agrupar todas as verifições de linha vazia
                HSSFSheet sheet = null;
                HSSFPatriarch drawingPatriarch = null;
                if (spreadsheetFactory.Header != null)
                {
                    sheet = CreateTitle(spreadsheetFactory.Header.Title, spreadsheetFactory.Name);
                    drawingPatriarch = sheet.CreateDrawingPatriarch();

                    if (spreadsheetFactory.Header.Filters != null && spreadsheetFactory.Header.Filters.Count > 0)
                    {
                        sheet = ConfigHeader(sheet, spreadsheetFactory.Header.Filters);
                    }

                    sheet.CreateRow(sheet.LastRowNum + 1);
                    sheet.CreateRow(sheet.LastRowNum + 1);

                    PrepareTableHeader(sheet, spreadsheetFactory, spreadsheetFactory.FirstHeaderCell);

                    int firstRow = (sheet.LastRowNum - spreadsheetFactory.TableHeaderRows);
                    ConfigTableHeader(sheet, spreadsheetFactory.TableHeaders, spreadsheetFactory.FirstHeaderCell, (sheet.LastRowNum - spreadsheetFactory.TableHeaderRows));

                    if (spreadsheetFactory.HeaderCellStyle != null)
                    {
                        ApplyHeaderCellStyle(sheet, firstRow, spreadsheetFactory.FirstHeaderCell, spreadsheetFactory.TableHeaderRows, spreadsheetFactory.TableHeaderCells, spreadsheetFactory.HeaderCellStyle);
                    }

                }
                else
                {
                    sheet = CreateTitle(null, spreadsheetFactory.Name);
                }

                if (spreadsheetFactory.Datasource != null)
                {
                    int firstRow = 0;
                    if (sheet.GetRow(sheet.LastRowNum).IsEmpty())
                    {
                        firstRow = sheet.LastRowNum;
                    }
                    else
                    {
                        firstRow = sheet.LastRowNum + 2;
                    }

                    //
                    if (!String.IsNullOrEmpty(spreadsheetFactory.MergedTitle))
                    {
                        //spreadsheetFactory.ChildSheet
                        CreateContentTableCells(sheet, spreadsheetFactory.Datasource.Count, spreadsheetFactory.Properties.Length, spreadsheetFactory.FirstCell, firstRow, spreadsheetFactory.MergedTitle);
                        //CreateContentTableCells(sheet, spreadsheetFactory, firstRow);
                        firstRow++;
                    }
                    else
                    {
                        CreateContentTableCells(sheet, spreadsheetFactory.Datasource.Count, spreadsheetFactory.Properties.Length, spreadsheetFactory.FirstCell, firstRow, string.Empty);
                        //CreateContentTableCells(sheet, spreadsheetFactory, firstRow);
                    }

                    

                    if (sheet.GetRow(sheet.LastRowNum).IsEmpty())
                    {
                        //TODO: corrigir
                        SetContentTableCellsValue(sheet, firstRow - 1, 0, spreadsheetFactory.Properties, spreadsheetFactory.Datasource,spreadsheetFactory.ChildSheet);
                    }
                    else
                    {
                        SetContentTableCellsValue(sheet, firstRow, 0, spreadsheetFactory.Properties, spreadsheetFactory.Datasource,spreadsheetFactory.ChildSheet);
                    }
                    
                    
                    if (spreadsheetFactory.ConditionalFormatList != null)
                    {
                        List<PropertyCell> teste = new List<PropertyCell>();
                        for (int i = 0; i < spreadsheetFactory.Properties.Length; i++)
                        {
                            if (spreadsheetFactory.ConditionalFormatList.Keys.Contains(spreadsheetFactory.Properties[i]))
                            {
                                teste.Add(new PropertyCell() { CellIndex = i, PropertyName = spreadsheetFactory.Properties[i] });
                            }
                        }

                        ApplyConditinalFormattingContentTable(sheet, firstRow, spreadsheetFactory.Datasource.Count, 0, spreadsheetFactory.Properties.Length + 0, teste, spreadsheetFactory.RowStyle, spreadsheetFactory.ConditionalFormatList);
                    }
                    //if (!String.IsNullOrEmpty(spreadsheetFactory.MergedTitle))
                    //{
                    //    sheet.CreateRow(sheet.LastRowNum);
                    //    sheet.GroupRow(firstRow, sheet.LastRowNum);
                    //    sheet.SetRowGroupCollapsed(firstRow, true);
                    //}
                    //else
                    //{
                        sheet.CreateRow(sheet.LastRowNum+1);
                        sheet.GroupRow(firstRow, sheet.LastRowNum-1);
                        //sheet.SetRowGroupCollapsed(firstRow, true);
                    //}
                }
            }
        }

        #endregion Public Methods

        #region Private Methods

        private void ApplyHeaderCellStyle(HSSFSheet sheet, int firstRow, int firstCell, int rows, int cells, HSSFCellStyle style)
        {
            for (int r = firstRow; r < (rows + firstRow); r++)
            {
                for (int c = firstCell; c < (cells + firstCell); c++)
                {
                    sheet.GetRow(r).GetCell(c).CellStyle = style;
                }
            }
        }

        private void ApplyConditinalFormattingContentTable(HSSFSheet sheet, int firstRow, int rows, int firstCell, int lastCell, List<PropertyCell> cellPoints, RowStyle defaultContentRowStyle, IDictionary<string, List<ConditionalFormattingTemplate>> conditionalFormatList)
        {
            for (int r = firstRow; r < (firstRow + rows); r++)
            {
                RowStyle style = GetRowStyle(sheet.GetRow(r), cellPoints, conditionalFormatList);

                if (style != null)
                {
                    ApplyRowStyle(sheet.GetRow(r), style);
                }
                else
                {
                    ApplyRowStyle(sheet.GetRow(r), defaultContentRowStyle);
                }
            }
        }

        private RowStyle GetRowStyle(HSSFRow row, List<PropertyCell> cellPoints, IDictionary<string, List<ConditionalFormattingTemplate>> conditionalFormatList)
        {
            ConditionalFormattingTemplate template = null;
            foreach (var item in cellPoints)
            {
                //TODO: testar quando nao tiver propriedade ou for nulo
                foreach (var style in conditionalFormatList[item.PropertyName])
                {
                    if (IsMatchStyle(style, row.GetCell(item.CellIndex)))
                    {
                        if (template != null)
                        {
                            if (style.Priority > template.Priority)
                            {
                                template = style;
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
                return template.RowStyle;
            }
            else
            {
                return null;
            }
        }

        private bool IsMatchStyle(ConditionalFormattingTemplate cft, HSSFCell cell)
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

        private void AddConditionalFormatComment(HSSFSheet sheet, HSSFCell cell, string comment)
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

        private void PrepareTableHeader(HSSFSheet sheet, SpreadsheetFactory spreadSheet, int cell)
        {
            CreateHeaderRows(sheet, spreadSheet.TableHeaders, spreadSheet.TableHeaderRows);
            CreateHeaderCells(sheet, cell, spreadSheet.TableHeaderRows, spreadSheet.TableHeaderCells);
        }

        private void ConfigTableHeader(HSSFSheet sheet, IList<TableHeader> tableHeaders, int cell, int row)
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
                        ConfigTableHeader(sheet, item.Cells, headerCell, cellRow + 1);
                    }
                }
            }
        }

        private void CreateHeaderRows(HSSFSheet sheet, IList<TableHeader> tableHeaders, int tableHeaderRows)
        {
            for (int i = 0; i < tableHeaderRows; i++)
            {
                sheet.CreateRow(sheet.LastRowNum + 1);
            }
        }

        private void CreateHeaderCells(HSSFSheet sheet, int firstCell, int tableHeaderRows, int tableHeaderCells)
        {
            for (int i = 1; i < tableHeaderRows + 1; i++)
            {
                for (int j = firstCell; j < tableHeaderCells + firstCell; j++)
                {
                    //sheet.GetRow(sheet.LastRowNum - i).CreateCell(j).SetCellValue(i + ":" + j);
                    sheet.GetRow(sheet.LastRowNum - i).CreateCell(j);
                }
                Console.WriteLine(sheet.LastRowNum - i);
            }
        }

        private void CreateContentTableCells(HSSFSheet sheet, int rows, int cells, int firstCell, int firstRow, string mergedTitle)
        {
            if (!String.IsNullOrEmpty(mergedTitle))
            {
                sheet.CreateRow(firstRow);
                for (int j = firstCell; j < firstCell + cells; j++)
                {
                    sheet.GetRow(firstRow).CreateCell(j);
                }
                sheet.GetRow(firstRow).GetCell(firstCell).SetCellValue(mergedTitle);
                sheet.AddMergedRegion(new CellRangeAddress(firstRow, firstRow, sheet.GetRow(firstRow).FirstCellNum, sheet.GetRow(firstRow).LastCellNum - 1));
                firstRow++;
            }

            for (int i = firstRow; i < firstRow + rows; i++)
            {
                sheet.CreateRow(i);
                for (int j = firstCell; j < firstCell + cells; j++)
                {
                    sheet.GetRow(i).CreateCell(j).SetCellValue("A");
                }
            }
        }

        private void CreateContentTableCells(HSSFSheet sheet, SpreadsheetFactory spreadsheet, int firstRow)
        {
            if (!String.IsNullOrEmpty(spreadsheet.MergedTitle))
            {
                sheet.CreateRow(firstRow);
                for (int j = spreadsheet.FirstCell; j < spreadsheet.FirstCell + spreadsheet.Properties.Length; j++)
                {
                    sheet.GetRow(firstRow).CreateCell(j);
                }
                sheet.GetRow(firstRow).GetCell(spreadsheet.FirstCell).SetCellValue(spreadsheet.MergedTitle);
                sheet.AddMergedRegion(new CellRangeAddress(firstRow, firstRow, sheet.GetRow(firstRow).FirstCellNum, sheet.GetRow(firstRow).LastCellNum - 1));
                firstRow++;
            }

            for (int i = firstRow; i < firstRow + spreadsheet.Datasource.Count; i++)
            {
                sheet.CreateRow(i);
                for (int j = spreadsheet.FirstCell; j < spreadsheet.FirstCell + spreadsheet.Properties.Length; j++)
                {
                    sheet.GetRow(i).CreateCell(j).SetCellValue("A");
                }
            }
        }

        private void SetContentTableCellsValue(HSSFSheet sheet, int firstRow, int firstCell, string[] properties, IList<object> values, ChildSheet child)
        {
            int row = firstRow;
            object propValue;

            if (child == null)
            {
                foreach (var item in values)
                {
                    for (int i = firstCell; i < properties.Length; i++)
                    {
                        propValue = GetPropValue(item, properties[i]);
                        switch (SheetUtil.GetCellType(propValue))
                        {
                            case HSSFCell.CELL_TYPE_NUMERIC:
                                sheet.GetRow(row).GetCell(i).SetCellValue(double.Parse(propValue.ToString()));
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
            else
            {
                IList<object> data = null;
                foreach (var item in values)
                {
                    child.Datasource = GetListValue(item, child.ChildPropertyName);

                    if (child.Datasource != null)
                    {
                        MountSpreadsheet(child);
                    }

                    for (int i = firstCell; i < properties.Length; i++)
                    {
                        propValue = GetPropValue(item, properties[i]);
                        switch (SheetUtil.GetCellType(propValue))
                        {
                            case HSSFCell.CELL_TYPE_NUMERIC:
                                sheet.GetRow(row).GetCell(i).SetCellValue(double.Parse(propValue.ToString()));
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
        }

        public static IList<object> GetListValue(object src, string propName)
        {
            //TODO: é realmente necessário utilizar uma lista generica de object?
            Type list = typeof(System.Collections.Generic.IList<object>);
            string listNamespace = list.Namespace;

            object aux = src;
            string[] nivels = propName.Split('.');

            object property = null;

            foreach (var nivel in nivels)
            {
                property = aux.GetType().GetProperty(nivel).GetValue(aux, null);
                if (property != null)
                {
                    aux = property;
                }
            }

            if (property != null && !property.GetType().Namespace.Equals(listNamespace))
            {
                //TODO: melhorar a mensagem informando a propriedade
                throw new ArgumentException("o objeto indicado  não é uma lista");
            }

            return property as List<object>;
        }

        private object GetPropValue(object src, string propName)
        {
            //return src.GetType().GetProperty(propName).GetValue(src, null);

            object aux = src;
            string[] nivels = propName.Split('.');
            object property = null;

            foreach (var nivel in nivels)
            {
                property = aux.GetType().GetProperty(nivel).GetValue(aux, null);
                if (property != null)
                {
                    aux = property;
                }
            }

            return property;
        }

        private HSSFSheet CreateTitle(string title, string sheetName)
        {
            HSSFSheet newSheet;

            if (!string.IsNullOrEmpty(sheetName))
            {
                newSheet = _workbook.GetSheet(sheetName);
                if (newSheet == null)
                {
                    newSheet = _workbook.CreateSheet(sheetName);
                }
            }
            else
            {
                newSheet = _workbook.GetSheet(sheetName);
                if (newSheet == null)
                {
                    newSheet = _workbook.CreateSheet(sheetName);
                }
            }

            if (!string.IsNullOrEmpty(title))
            {
                newSheet.CreateRow(newSheet.LastRowNum).CreateCell(0).SetCellValue(title);
                newSheet.AddMergedRegion(new CellRangeAddress(newSheet.LastRowNum, newSheet.LastRowNum, 0, SheetUtil.DEFAULT_TITLE_SPAN_SIZE));
            }

            return newSheet;
        }

        private HSSFSheet ConfigHeader(HSSFSheet sheet, IDictionary<string, object> filters)
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

        public void SaveSpreadsheet(string filePath, string fileName)
        {
            MemoryStream ms = new MemoryStream();
            _workbook.Write(ms);
            FileStream fs = new FileStream(filePath + fileName, FileMode.Create);
            fs.Write(ms.ToArray(), 0, Convert.ToInt32(ms.Length));
            fs.Close();
            ms.Close();

            _workbook = null;
        }

        private void ApplyRowStyle(HSSFRow row, RowStyle rowStyle)
        {
            if (rowStyle != null && rowStyle.CellStyles != null)
            {
                foreach (KeyValuePair<int, HSSFCellStyle> style in rowStyle.CellStyles)
                {
                    row.GetCell(style.Key + row.FirstCellNum).CellStyle = style.Value;
                }
            }
        }

        #endregion Private Methods

        #region Constructor

        public WorkbookManager()
        {
            if (_workbook == null)
            {
                _workbook = new HSSFWorkbook();
            }
        }

        #endregion Constructor

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
        public IList<object> Lista { get; set; }
    }
}
