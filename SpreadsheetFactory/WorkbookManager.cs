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

        private static int tableHeaderCells = 0;

        private static int headerCell;

        private static int tableHeaderCellsAux = 0;

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

        private static IList<object> contentList;

        private static IDictionary<string, string> contentTableTitleList;

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

                    sheet.CreateRow(sheet.LastRowNum + 1);
                    sheet.CreateRow(sheet.LastRowNum + 1);

                    PrepareTableHeader(sheet, spreadsheetFactory.TableHeaders, 0);
                    int cellAux = 0;
                    ConfigTableHeader(sheet, spreadsheetFactory.TableHeaders, ref cellAux, (sheet.LastRowNum - tableHeaderRows));
                    List<object> list = new List<object>();
                    list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "Italo", Salario = 10.3 });
                    list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "Italo", Salario = 10.3 });
                    list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "Italo", Salario = 10.3 });
                    list.Add(new Pessoa() { Idade = 12, Nascimento = DateTime.Now, Nome = "Italo", Salario = 10.3 });

                    string[] properties = new string[4];
                    properties[0] = "Idade";
                    properties[1] = "Nascimento";
                    properties[2] = "Nome";
                    properties[3] = "Salario";

                    int firstRow = sheet.LastRowNum;
                    CreateContentTableCells(sheet, list.Count, list.Count, 0, firstRow);
                    SetContentTableCellsValue(sheet, firstRow, 0, properties, list);
                }
            }
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
                    sheet.GetRow(cellRow)
                        .GetCell(headerCell)
                        .SetCellValue(item.Text);

                    //mesclar linhas
                    sheet.AddMergedRegion(new CellRangeAddress(cellRow, sheet.LastRowNum - 1, headerCell, headerCell + item.SpanSize));
                    headerCell++;
                }
                else if (item.Cells != null)
                {
                    sheet.GetRow(cellRow)
                        .GetCell(headerCell)
                        .SetCellValue(item.Text);

                    int headerCellAux = headerCell + item.Cells.Count - 1;

                    //mesclar celulas
                    sheet.AddMergedRegion(new CellRangeAddress(cellRow, cellRow, headerCell, headerCell + item.SpanSize));

                    bool callRecursive = true;
                    foreach (var internItem in item.Cells)
                    {
                        if (internItem.Cells == null || internItem.Cells.Count == 0)
                        {
                            sheet.GetRow(cellRow + 1)
                                .GetCell(headerCell)
                                .SetCellValue(internItem.Text);
                            //mesclar linhas
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
            for (int i = tableHeaderRows + 1; i > 0; i--)
            {
                for (int j = firstCell; j < tableHeaderCells + firstCell; j++)
                {
                    //sheet.GetRow(sheet.LastRowNum - i).CreateCell(j).SetCellValue(i + ":" + j);
                    sheet.GetRow(sheet.LastRowNum - i).CreateCell(j);
                }
            }
        }

        private static void CreateContentTableCells(HSSFSheet sheet, int rows, int cells, int firstCell, int firstRow)
        {
            for (int i = sheet.LastRowNum; i < firstRow + rows; i++)
            {
                for (int j = firstCell; j < firstCell + cells; j++)
                {
                    sheet.CreateRow(i).CreateCell(j).SetCellValue("A");
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
                        case HSSFCell.CELL_TYPE_NUMERIC:
                            sheet.GetRow(row).GetCell(i).SetCellValue(double.Parse(propValue.ToString()));
                            break;
                        case HSSFCell.CELL_TYPE_STRING:
                            sheet.GetRow(row).GetCell(i).SetCellValue(propValue.ToString());
                            break;
                        case SheetUtil.CELL_TYPE_DATETIME:
                            sheet.GetRow(row).GetCell(i).SetCellValue(((DateTime)propValue).ToShortDateString());
                            break;
                    }
                    //sheet.GetRow(row).GetCell(i).SetCellValue("A");
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
                    tableHeaderCells++;
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
