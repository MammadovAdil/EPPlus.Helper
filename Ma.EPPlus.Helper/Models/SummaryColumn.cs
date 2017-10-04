using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ma.EPPlus.Helper.Models
{
    /// <summary>
    /// Summarization detail of the value field.
    /// </summary>
    public class SummaryColumn
    {
        /// <summary>
        /// Summarized value field name.
        /// </summary>
        public string FieldName { get; set; }

        /// <summary>
        /// Name of the column in the pivot table.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Type of the calculation.
        /// </summary>
        public DataFieldFunctions DataFieldFunction { get; set; } = DataFieldFunctions.Sum;

        /// <summary>
        /// Creates instance of SummaryColumn.
        /// </summary>
        /// <param name="fieldName">Summarized value field name.</param>
        public SummaryColumn(string fieldName)
        {
            FieldName = fieldName;
            Name = fieldName;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fieldName">Summarized value field name.</param>
        /// <param name="name">Name of the column in the pivot table.</param>
        public SummaryColumn(string fieldName, string name)
            : this(fieldName)
        {
            Name = name;
        }

        /// <summary>
        ///  Creates instance of SummaryColumn.
        /// </summary>
        /// <param name="fieldName">Summarized value field name.</param>
        /// <param name="dataFieldFunction">Type of the calculation.</param>
        public SummaryColumn(string fieldName, DataFieldFunctions dataFieldFunction)
            : this(fieldName)
        {
            DataFieldFunction = dataFieldFunction;
        }

        /// <summary>
        ///  Creates instance of SummaryColumn.
        /// </summary>
        /// <param name="fieldName">Summarized value field name.</param>
        /// <param name="name">Name of the column in the pivot table.</param>
        /// <param name="dataFieldFunction">Type of the calculation.</param>
        public SummaryColumn(string fieldName, string name, DataFieldFunctions dataFieldFunction)
          : this(fieldName, name)
        {
            DataFieldFunction = dataFieldFunction;
        }
    }
}
