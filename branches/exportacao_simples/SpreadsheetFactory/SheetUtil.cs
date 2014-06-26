using System;
using System.Collections.Generic;
using System.Text;

namespace SpreadsheetFactory
{
    public static class SheetUtil
    {
        public static int DEFAULT_TITLE_SPAN_SIZE = 10;
        public static int DEFAULT_FIRST_CELL = 0;
        private static int tableHeaderRows = 0;
        private static int tableHeaderRowsAux = 0;

        //temporario, a biblioteca padrao nao possui um valor para Datetime
        public const int CELL_TYPE_DATETIME = 99;

        public static int GetCellType(object value)
        {
            Type type = value.GetType();

            switch (type.FullName)
            {
                case "System.Int16":
                case "System.Int32":
                case "System.Int64":
                case "System.Double":
                    return NPOI.HSSF.UserModel.HSSFCell.CELL_TYPE_NUMERIC;
                case "System.String":
                    return NPOI.HSSF.UserModel.HSSFCell.CELL_TYPE_STRING;
                case "System.DateTime":
                    return SheetUtil.CELL_TYPE_DATETIME;
                default:
                    return NPOI.HSSF.UserModel.HSSFCell.CELL_TYPE_STRING;
            }
        }

        public static int GetTotalHeaderRows(IList<TableHeader> tableHeaders)
        {
            foreach (var item in tableHeaders)
            {
                if (item.Cells == null)
                {
                    tableHeaderRowsAux++;
                    if (tableHeaderRowsAux > tableHeaderRows)
                    {
                        tableHeaderRows = tableHeaderRowsAux;
                    }
                }
                else if (item.Cells != null)
                {
                    tableHeaderRowsAux++;
                    bool callRecursive = true;
                    foreach (var internItem in item.Cells)
                    {
                        if (internItem.Cells == null || internItem.Cells.Count == 0)
                        {
                            callRecursive = false;
                        }
                    }

                    if (callRecursive)
                    {
                        //teste++;
                        GetTotalHeaderRows(item.Cells);
                        if (tableHeaderRowsAux > tableHeaderRows)
                        {
                            tableHeaderRows = tableHeaderRowsAux;
                        }
                    }
                    else
                    {
                        tableHeaderRowsAux++;
                        if (tableHeaderRowsAux > tableHeaderRows)
                        {
                            tableHeaderRows = tableHeaderRowsAux;
                        }
                    }
                }
                if (tableHeaderRowsAux > tableHeaderRows)
                {
                    tableHeaderRows = tableHeaderRowsAux;
                }
                tableHeaderRowsAux = 0;
            }
            return tableHeaderRows;
        }
    }
}
