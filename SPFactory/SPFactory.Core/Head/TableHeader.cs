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
    /// Class to represent the principal table header.
    /// </summary>
    public class TableHeader
    {
        #region Private Members

        private string _text;
        private IList<TableHeader> _cells = null;

        #endregion

        #region Public Properties

        /// <summary>
        /// Cell value.
        /// </summary>
        public string Text
        {
            get { return _text; }
            set { _text = value; }
        }

        /// <summary>
        /// Used to group TableHeader values.
        /// </summary>
        public IList<TableHeader> Cells
        {
            get { return _cells; }
            set { _cells = value; }
        }

        /// <summary>
        /// Number of cells to span.
        /// </summary>
        public int SpanSize
        {
            get
            {
                if (_cells != null)
                {
                    return CellSpanCount(_cells) - 1;
                }
                else
                {
                    return 0;
                }
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public TableHeader()
        {

        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Add a child cell to the current header.
        /// </summary>
        /// <param name="value">Child cell text.</param>
        /// <returns>A new TableHeader</returns>
        public TableHeader AddSpanCell(string value)
        {
            TableHeader tableHeader = new TableHeader();
            tableHeader.Text = value;
            if (_cells == null)
            {
                _cells = new List<TableHeader>();
            }
            _cells.Add(tableHeader);


            return tableHeader;
        }

        /// <summary>
        /// Count the number of cells to span.
        /// </summary>
        /// <param name="list">List of grouped Table Headers</param>
        /// <returns>Total of cell span.</returns>
        private int CellSpanCount(IList<TableHeader> list)
        {
            int total = 0;
            foreach (var item in list)
            {
                if (item.Cells == null)
                {
                    total++;
                }
                else
                {
                    total += CellSpanCount(item.Cells);
                }
            }

            return total;
        }

        #endregion
    }
}
