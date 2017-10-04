using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Ma.EPPlus.Helper
{
    /// <summary>
    /// Helper methods to easily read Excel files.
    /// </summary>
    public static class ReadingHelpers
    {
        /// <summary>
        /// Convert excel worksheet to DataTable.
        /// </summary>
        /// <exception cref="ArgumentNullException">
        /// When worksheet is null.
        /// </exception>
        /// <param name="worksheet">Worksheet to convert.</param>
        /// <param name="hasHeaders">If worksheet has headers.</param>
        /// <returns>DataTable.</returns>
        public static DataTable ToDataTable(
            this ExcelWorksheet worksheet,
            bool hasHeaders)
        {
            DataTable dataTable = new DataTable();

            /// Add columns. If worksheet has headers then get them from worksheet.
            /// Otherwise add empty columns.
            if (hasHeaders)
            {
                string pattern = "[^a-zA-Z0-9]";
                Regex unwantedChars = new Regex(pattern);
                foreach (ExcelRangeBase firstRowCell
                    in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    // Remove unwanted chars from column header
                    string columnHeader = firstRowCell.Text.ToLower(CultureInfo.InvariantCulture);
                    columnHeader = unwantedChars.Replace(columnHeader, string.Empty);
                    dataTable.Columns.Add(columnHeader);
                }
            }
            else
            {
                for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
                {
                    dataTable.Columns.Add();
                }
            }

            // Define data start row
            int startRowIndex = hasHeaders ? 2 : 1;
            for (int rowNumber = startRowIndex; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
            {
                ExcelRange row = worksheet
                    .Cells[rowNumber, 1, rowNumber, dataTable.Columns.Count];
                DataRow dataTableRow = dataTable.NewRow();
                foreach (ExcelRangeBase cell in row)
                {
                    dataTableRow[cell.Start.Column - 1] = cell.Text;
                }
                dataTable.Rows.Add(dataTableRow);
            }
            return dataTable;
        }

        /// <summary>
        /// Cast excel worksheet to list of model. Workseet must have headers.
        /// Proeprties are matched using column headers. 
        /// Case and special carecters like (-, _) are ignored.
        /// </summary>
        /// <exception cref="ArgumentNullException">
        /// When worksheet is null.
        /// </exception>
        /// <typeparam name="TModel">Type of model.</typeparam>
        /// <param name="worksheet">Worksheet to cast.</param>
        /// <returns>List of model.</returns>
        public static List<TModel> Cast<TModel>(this ExcelWorksheet worksheet)
            where TModel : class, new()
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            // Cast ot dataTable
            DataTable dataTable = ToDataTable(worksheet, true);

            // Find and remove empty rows
            List<DataRow> nullRows = dataTable
                .AsEnumerable()
                .Where(m => m.ItemArray.All(item => item == null
                    || item is DBNull
                    || item == DBNull.Value
                    || item.Equals(string.Empty)))
                .ToList();
            nullRows.ForEach(m => dataTable.Rows.Remove(m));

            return Cast<TModel>(dataTable);
        }

        /// <summary>
        /// Cast data table to list of model.DataTable must have column names.
        /// Proeprties are matched using column names. 
        /// Case and special carecters like (-, _) are ignored.
        /// </summary>
        /// <exception cref="ArgumentNullException">
        /// When dataTable is null.
        /// </exception>
        /// <typeparam name="TModel">Type of model.</typeparam>
        /// <param name="dataTable">DataTable to cast.</param>
        /// <returns>List of models.</returns>
        public static List<TModel> Cast<TModel>(this DataTable dataTable)
            where TModel : class, new()
        {
            if (dataTable == null)
                throw new ArgumentNullException(nameof(dataTable));

            // Initialize model list
            List<TModel> modelList = new List<TModel>();

            // Get properties of model
            List<PropertyInfo> properties = typeof(TModel)
                .GetProperties()
                .ToList();

            // Filter proeprties and remove those which no column exist in dataTable.
            properties.RemoveAll(m => !dataTable.Columns.Contains(
                m.Name.ToLower(CultureInfo.InvariantCulture)));

            // Read data from data table and set properties of model.
            foreach (DataRow row in dataTable.AsEnumerable())
            {
                // Initialize model
                TModel model = new TModel();
                foreach (PropertyInfo property in properties)
                {
                    // Read cell value from row
                    object cellValue = row[property.Name.ToLower(CultureInfo.InvariantCulture)];

                    // Consider DBNull
                    if (cellValue == DBNull.Value)
                        cellValue = null;

                    // Convert to needed type
                    if (cellValue != null)
                    {
                        // For nullable types
                        Type underLyingType = Nullable.GetUnderlyingType(property.PropertyType);
                        Type propertyType = underLyingType ?? property.PropertyType;

                        cellValue = Convert.ChangeType(cellValue, propertyType);
                    }

                    // Assign proeprty value
                    property.SetValue(model, cellValue);
                }
                modelList.Add(model);
            }

            return modelList;
        }
    }
}
