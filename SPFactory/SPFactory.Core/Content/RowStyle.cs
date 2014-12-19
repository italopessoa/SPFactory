using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace SPFactory.Core.Content
{
    /// <summary>
    /// Class to mount the row style template.
    /// </summary>
    public class RowStyle
    {
        #region Private Members

        private IDictionary<int, NPOI.HSSF.UserModel.HSSFCellStyle> _cellStyles;
        private string _rowStyleName;
        private short? _rowForegroundColorIndex;
        private short? _rowFillPattern;

        #endregion

        #region Public Properties

        /// <summary>
        /// Foreground color value.
        /// </summary>
        public short? RowForegroundColor
        {
            get { return _rowForegroundColorIndex; }
            set { _rowForegroundColorIndex = value; }
        }

        /// <summary>
        /// Fill pattern value.
        /// </summary>
        public short? RowFillPattern
        {
            get { return _rowFillPattern; }
            set { _rowFillPattern = value; }
        }

        /// <summary>
        /// Cell style dictionary.
        /// </summary>
        public IDictionary<int, HSSFCellStyle> CellStyles
        {
            get { return _cellStyles; }
            private set { _cellStyles = value; }
        }

        /// <summary>
        /// Row style name.
        /// </summary>
        public string RowStyleName
        {
            get { return _rowStyleName; }
            set { _rowStyleName = value; }
        }

        #endregion Properties

        #region Constructors
        
        /// <summary>
        /// Empty constructor.
        /// </summary>
        public RowStyle()
        {
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellIndex"></param>
        /// <param name="cellStyle"></param>
        public void AddCellRowStyle(int cellIndex, HSSFCellStyle cellStyle)
        {
            if (_cellStyles == null)
            {
                _cellStyles = new Dictionary<int, NPOI.HSSF.UserModel.HSSFCellStyle>();
            }

            if (_rowForegroundColorIndex.HasValue)
            {
                cellStyle.FillForegroundColor = _rowForegroundColorIndex.Value;
            }

            if (_rowFillPattern.HasValue)
            {
                cellStyle.FillPattern = _rowFillPattern.Value;
            }

            if (_cellStyles.Keys.Contains(cellIndex))
            {
                _cellStyles[cellIndex] = cellStyle;
            }
            else
            {
                _cellStyles.Add(cellIndex, cellStyle);
            }
        }

        #endregion
    }
}
