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
using SPFactory.Core.Factory;
using SPFactory.Core.Head;
using SPFactory.Core.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace SPFactory.Core.Content
{
    /// <summary>
    /// Class to join all Spreadsheet components
    /// </summary>
    public class Spreadsheet
    {
        #region Private Members

        private Header _header;
        private IList<TableHeader> _tableHeaders;
        private IList<object> _datasource;
        private string[] _properties;
        private IList<Spreadsheet> _spreadsheetList;
        private int _firstCell = 0;
        private HSSFCellStyle _spanTitleStyle;
        private ChildSheet _childSheet;
        
        #endregion

        #region Public Properties

        /// <summary>
        /// Sheet filter header.
        /// </summary>
        public Header Header
        {
            get { return _header; }
            set { _header = value; }
        }

        /// <summary>
        /// Principal table header.
        /// </summary>
        public IList<TableHeader> TableHeaders
        {
            get { return _tableHeaders; }
            set { _tableHeaders = value; }
        }

        /// <summary>
        /// List of objects to bind bind with the sheet.
        /// </summary>
        public IList<object> Datasource
        {
            get { return _datasource; }
            set { _datasource = value; }
        }

        /// <summary>
        /// List of class's properties used to get the value of each Datasource object.
        /// </summary>
        public string[] Properties
        {
            get { return _properties; }
            set { _properties = value; }
        }

        /// <summary>
        /// Number of rows of principal table header (TableHeader).
        /// </summary>
        public int TableHeaderRows
        {
            get { return SheetUtil.GetTotalHeaderRows(_tableHeaders); }
        }

        /// <summary>
        /// Number of cells of principal table header (TableHeader).
        /// </summary>
        public int TableHeaderCells
        {
            get { return SheetUtil.GetTableHeaderCells(_tableHeaders); }
        }

        /// <summary>
        /// Firts cell of principal table header (TableHeader).
        /// </summary>
        public int FirstHeaderCell { get; set; }

        /// <summary>
        /// List of Spreadsheets (tabs of Excel/Calc).
        /// </summary>
        public IList<Spreadsheet> SpreadsheetList
        {
            get { return _spreadsheetList; }
            set { _spreadsheetList = value; }
        }

        /// <summary>
        /// Sheet's name (the value displayed at the Excel/Calc tabs).
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Text displayed at the first merged line (Title).
        /// </summary>
        public string MergedTitle { get; set; }

        /// <summary>
        /// Style and format to apply in the header cells.
        /// A unique style is applied to all the cells.
        /// </summary>
        public HSSFCellStyle HeaderCellStyle { get; set; }

        /// <summary>
        /// The firts cell of content.
        /// </summary>
        public int FirstCell
        {
            get { return _firstCell; }
            set { _firstCell = value; }
        }

        /// <summary>
        /// Style and format to apply in the content title.
        /// </summary>
        public HSSFCellStyle SpanTitleStyle
        {
            get { return _spanTitleStyle; }
            set { _spanTitleStyle = value; }
        }

        #endregion

        #region Constructors
        
        /// <summary>
        /// Default constructor
        /// </summary>
        public Spreadsheet()
        {

        }

        #endregion

        #region ainda não testado

        public virtual RowStyle RowStyle { get; set; }

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

        #endregion

        //private SpreadsheetFactory _childSheet;
        //private ChildSheet _childSheet;

        public ChildSheet ChildSheet
        {
            get { return _childSheet; }
            set { _childSheet = value; }
        }
    }
}
