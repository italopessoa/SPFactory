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

using System;
using System.Collections.Generic;
using System.Text;

namespace SPFactory.Core.Head
{
    /// <summary>
    /// Class to represent the header of the spreadsheet.
    /// Here you put informations like title, date, filters, etc.
    /// </summary>
    public class Header
    {
        #region Private Members

        private int DEFAULT_TITLE_SPAN_SIZE = 10;

        private IDictionary<string, object> _filters;
        private string _title;
        private string _sheetName;

        #endregion

        #region Public Properties

        /// <summary>
        /// Values used to get information for generate the spreadsheet
        /// </summary>
        public IDictionary<string, object> Filters
        {
            get { return _filters; }
        }

        /// <summary>
        /// Spreasheet title
        /// </summary>
        public string Title
        {
            get { return _title; }
            set { _title = value; }
        }

        /// <summary>
        /// Sheet name. It is important if you have more than one sheet.
        /// </summary>
        public string SheetName
        {
            get { return _sheetName; }
            set { _sheetName = value; }
        }

        /// <summary>
        /// The number of columns to the sheet span.
        /// </summary>
        private int TitleSpan
        {
            get { return DEFAULT_TITLE_SPAN_SIZE; }
            set { DEFAULT_TITLE_SPAN_SIZE = value; }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public Header()
        {

        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Add new filters to the sheet.
        /// </summary>
        /// <param name="key">Value used to identify a filter, i.e. "name", "birth date", "id"...</param>
        /// <param name="value">Filter's value.</param>
        /// <returns>Return true if the value was successfully inserted.</returns>
        public bool AddFilter(string key, object value)
        {
            if (ValidadeHeaderValues(key, value))
            {
                if (_filters == null)
                {
                    _filters = new Dictionary<string, object>();
                }

                if (!_filters.Keys.Contains(key))
                {
                    _filters.Add(key, value);
                }

                return true;
            }
            return false;
        }

        /// <summary>
        /// Verify a value before insert it in the filters dictionary.
        /// </summary>
        /// <param name="key">Filter's name.</param>
        /// <param name="value">Filter's value</param>
        /// <returns>Return true the value can be inserted, and false otherwise.</returns>
        private bool ValidadeHeaderValues(string key, object value)
        {
            //TODO: implement filter value verification
            return true;
        }

        #endregion        
    }
}
