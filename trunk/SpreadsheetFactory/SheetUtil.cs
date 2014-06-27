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
        private static int _tableHeaderCells = 0;
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

        public static int GetTableHeaderCells(IList<TableHeader> tableHeaders)
        {
            _tableHeaderCells = 0;
            CountTableHeaderCells(tableHeaders);
            return _tableHeaderCells;
        }

        private static void CountTableHeaderCells(IList<TableHeader> tableHeaders)
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
                    CountTableHeaderCells(item.Cells);
                }
            }
        }

        //TODO: melhorar essa verificação, uma célula pode não ter um tip definido. se for bool sempre terá um valor. Linha vazia não contém células
        public static bool IsEmpty(this NPOI.HSSF.UserModel.HSSFRow row)
        {
            return row.FirstCellNum < 0;
        }
    }
}

// you need this once (only), and it must be in this namespace
namespace System.Runtime.CompilerServices
{
    [AttributeUsage(AttributeTargets.Assembly | AttributeTargets.Class | AttributeTargets.Method)]
    public sealed class ExtensionAttribute : Attribute { }
}
