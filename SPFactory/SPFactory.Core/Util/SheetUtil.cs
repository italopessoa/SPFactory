/*
 * SPFactory is a class library to generate custom spreadsheets using NPOI.
 * Copyright (C) 2014 Italo Pessoa (italo.pessoa@hotmail.com)
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

using NPOI.HSSF.UserModel;
using SPFactory.Core.Head;
using System;
using System.Collections.Generic;
using System.Text;

namespace SPFactory.Core.Util
{
    /// <summary>
    /// Utilities methods and properties.
    /// </summary>
    public static class SheetUtil
    {

        #region Private Members

        /// <summary>
        /// First cell location
        /// </summary>
        public static int DEFAULT_FIRST_CELL = 0;
        private static int tableHeaderRows = 0;
        private static int tableHeaderRowsAux = 0;
        private static int _tableHeaderCells = 0;

        #endregion

        #region Public Members

        /// <summary>
        /// Number of cells to span
        /// </summary>
        public static int DEFAULT_TITLE_SPAN_SIZE = 10;

        //temporario, a biblioteca padrao nao possui um valor para Datetime
        /// <summary>
        /// Datetime cell type identifier
        /// </summary>
        public const int CELL_TYPE_DATETIME = 99;

        #endregion

        #region Public Methods

        /// <summary>
        /// Identify the cell's value type.
        /// </summary>
        /// <param name="value">Cell value</param>
        /// <returns>Value type id.</returns>
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

        /// <summary>
        /// Number of rows of Table Header.
        /// </summary>
        /// <param name="tableHeaders">Table Header</param>
        /// <returns>Number of rows</returns>
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

        /// <summary>
        /// Number of cells of Table Header.
        /// </summary>
        /// <param name="tableHeaders">Table Header</param>
        /// <returns>Number of cells</returns>
        public static int GetTableHeaderCells(IList<TableHeader> tableHeaders)
        {
            _tableHeaderCells = 0;
            CountTableHeaderCells(tableHeaders);
            return _tableHeaderCells;
        }

        //TODO: melhorar essa verificação, uma célula pode não ter um tip definido. se for bool sempre terá um valor. Linha vazia não contém células
        /// <summary>
        /// Extension method to verify if a row is empty
        /// </summary>
        /// <param name="row">Spreasheet row</param>
        /// <returns>True if is empty false otherwise.</returns>
        public static bool IsEmpty(this HSSFRow row)
        {
            return row.FirstCellNum < 0;
        }

        #endregion

        #region Private Methods
        
        /// <summary>
        /// Number of cells of Table Header.
        /// </summary>
        /// <param name="tableHeaders">Table Header</param>
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

        #endregion
    }
}

// you need this once (only), and it must be in this namespace
namespace System.Runtime.CompilerServices
{
    /// <summary>
    /// ExtensionAttribute
    /// </summary>
    [AttributeUsage(AttributeTargets.Assembly | AttributeTargets.Class | AttributeTargets.Method)]
    public sealed class ExtensionAttribute : Attribute { }
}