using SPFactory.Core.Factory;
using System;
using System.Collections.Generic;
using System.Text;

namespace SPFactory.Core.Content
{
    /// <summary>
    /// 
    /// </summary>
    public class ChildSheet : Spreadsheet
    {
        #region Private Members

        private string _childPropertyName;

        #endregion

        #region Public Properties

        /// <summary>
        /// Spreadsheet name.
        /// </summary>
        public string ChildPropertyName
        {
            get { return _childPropertyName; }
        }

        #endregion

        #region Constructors
        
        /// <summary>
        /// Constructor to initialize a new ChildSheet.
        /// </summary>
        /// <param name="property">The name of the property that contains the ChildSheet data.</param>
        public ChildSheet(string property)
        {
            _childPropertyName = property;
        }

        #endregion

        #region Private Methods

        private IDictionary<string, List<ConditionalFormattingTemplate>> _conditionalFormatDictionary;

        #endregion

        #region Public Methods

        public override IDictionary<string, List<ConditionalFormattingTemplate>> ConditionalFormatList
        {
            get { return _conditionalFormatDictionary; }
        }

        public override void AddConditionalFormatting(string property, ConditionalFormattingTemplate format)
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
        //public override RowStyle RowStyle { get; set; }
    }
}
