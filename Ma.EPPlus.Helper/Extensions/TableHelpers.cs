using Ma.EPPlus.Helper.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Ma.EPPlus.Helper.Extensions
{
    /// <summary>
    /// Helper methods to manage excel tables.
    /// </summary>
    public static class TableHelpers
    {
        /// <summary>
        /// Set address of existing Excel table.
        /// </summary>
        /// <exception cref="ArgumentNullException">
        /// When table, startAddress or endAddress is null.
        /// </exception>
        /// <remarks>
        /// Currently it is not possible thorugh EPPlus properties,
        /// as they are read only.
        /// </remarks>
        /// <param name="table">Excel table to set address.</param>
        /// <param name="startAddress">Start address of table.</param>
        /// <param name="endAddress">End address of table.</param>
        public static void SetAddress(
            this ExcelTable table,
            ExcelCellAddress startAddress,
            ExcelCellAddress endAddress)
        {
            if (table == null)
                throw new ArgumentNullException(nameof(table));
            if (startAddress == null)
                throw new ArgumentNullException(nameof(startAddress));
            if (startAddress == null)
                throw new ArgumentNullException(nameof(endAddress));

            string tableRange = string.Format("{0}:{1}", startAddress.Address, endAddress.Address);
            var tableXml = table.TableXml.DocumentElement;
            tableXml.Attributes["ref"].Value = tableRange;
            tableXml["autoFilter"].Attributes["ref"].Value = tableRange;
        }

        /// <summary>
        /// Add Pivot table according to excel table..
        /// </summary>
        /// <remarks>
        /// Slightly modified version of
        /// https://stackoverflow.com/a/13979855/1380428.
        /// </remarks>
        /// <exception cref="ArgumentNullException">
        /// When package or table is null.
        /// </exception>
        /// <exception cref="ArgumentException">
        /// When no group by or summary column has been provided.
        /// </exception>
        /// <param name="package">Excel package to add pivot table to.</param>
        /// <param name="table">Table to add pivot table for.</param>
        /// <param name="groupByColumns">Columns to group according to.</param>
        /// <param name="summaryColumns">Columns to show summary for.</param>
        /// <returns>Added pivot table.</returns>
        public static ExcelPivotTable AddPivotTable(
            this ExcelPackage package,
            ExcelTable table,
            List<string> groupByColumns,
            List<SummaryColumn> summaryColumns,
            string pivotWorksheetName = null)
        {
            return package.AddPivotTable(table, groupByColumns, summaryColumns, null, pivotWorksheetName);
        }

        ///// <summary>
        ///// Add Pivot table according to excel table..
        ///// </summary>
        ///// <remarks>
        ///// Slightly modified version of
        ///// https://stackoverflow.com/a/13979855/1380428.
        ///// </remarks>
        ///// <exception cref="ArgumentNullException">
        ///// When package or table is null.
        ///// </exception>
        ///// <exception cref="ArgumentException">
        ///// When no group by or summary column has been provided.
        ///// </exception>
        ///// <param name="package">Excel package to add pivot table to.</param>
        ///// <param name="table">Table to add pivot table for.</param>
        ///// <param name="groupByColumns">Columns to group according to.</param>
        ///// <param name="summaryColumns">Columns to show summary for.</param>
        ///// <param name="filterColumns">Columns to add filter for.</param>
        ///// <returns>Added pivot table.</returns>
        //public static ExcelPivotTable AddPivotTable(
        //    this ExcelPackage package,
        //    ExcelTable table,
        //    List<string> groupByColumns,
        //    List<FJSummaryColumn> summaryColumns,
        //    List<string> filterColumns)
        //{
        //    return package.AddPivotTable(table, groupByColumns, summaryColumns, filterColumns, null);
        //}

        /// <summary>
        /// Add Pivot table according to excel table.
        /// </summary>
        /// <remarks>
        /// Slightly modified version of
        /// https://stackoverflow.com/a/13979855/1380428.
        /// </remarks>
        /// <exception cref="ArgumentNullException">
        /// When package or table is null.
        /// </exception>
        /// <exception cref="ArgumentException">
        /// When no group by or summary column has been provided.
        /// </exception>
        /// <param name="package">Excel package to add pivot table to.</param>
        /// <param name="table">Table to add pivot table for.</param>
        /// <param name="groupByColumns">Columns to group according to.</param>
        /// <param name="summaryColumns">Columns to show summary for.</param>
        /// <param name="filterColumns">Columns to add filter for.</param>
        /// <param name="pivotWorksheetName">Name of pivot worksheet.</param>
        /// <returns>Added pivot table.</returns>
        public static ExcelPivotTable AddPivotTable(
            this ExcelPackage package,
            ExcelTable table,
            List<string> groupByColumns,
            List<SummaryColumn> summaryColumns,
            List<string> filterColumns,
            string pivotWorksheetName = null)
        {
            if (package == null)
                throw new ArgumentNullException(nameof(package));
            if (table == null)
                throw new ArgumentNullException(nameof(table));
            if (groupByColumns == null || groupByColumns.Count == 0)
                throw new ArgumentException("At least one group by column must be provided.");
            if (summaryColumns == null || summaryColumns.Count == 0)
                throw new ArgumentException("At least one summary column must be provided.");
            if(summaryColumns.Any(c=> String.IsNullOrEmpty(c.FieldName)))
                throw new ArgumentException("Field name can not be empty.");

            // Initialize workseet name if not set.
            if (string.IsNullOrEmpty(pivotWorksheetName))
                pivotWorksheetName = "Pivot-" + table.Name.Replace(" ", "");
            var wsPivot = package.Workbook.Worksheets.Add(pivotWorksheetName);

            // Define pivot start row according to filter fileds
            int pivotStartRow = 1;
            if (filterColumns != null)
                pivotStartRow += filterColumns.Count;
            ExcelRange pivotStartAddress = wsPivot.Cells[pivotStartRow, 1];

            var dataRange = table.WorkSheet.Cells[table.Address.ToString()];
            var pivotTable = wsPivot.PivotTables.Add(
                pivotStartAddress,
                dataRange,
                "Pivot_" + table.Name.Replace(" ", ""));

            pivotTable.ShowHeaders = true;
            pivotTable.UseAutoFormatting = true;
            pivotTable.ApplyWidthHeightFormats = true;
            pivotTable.ShowDrill = true;
            pivotTable.FirstHeaderRow = 1;  // first row has headers
            pivotTable.FirstDataCol = 1;    // first col of data
            pivotTable.FirstDataRow = 2;    // first row of data

            // Add group by columns
            foreach (string row in groupByColumns)
            {
                var field = pivotTable.Fields[row];
                pivotTable.RowFields.Add(field);
                field.Sort = eSortType.Ascending;
            }

            // Add summary columns
            foreach (var column in summaryColumns)
            {
                var field = pivotTable.DataFields.Add(pivotTable.Fields[column.FieldName]);

                field.Function = column.DataFieldFunction;
                field.Name = column.Name;
            }

            // Add filter columns
            if (filterColumns != null)
            {
                foreach (string filter in filterColumns)
                {
                    var field = pivotTable.Fields[filter];
                    field.Sort = eSortType.Ascending;
                    pivotTable.PageFields.Add(field);
                }
            }

            pivotTable.DataOnRows = false;
            return pivotTable;
        }
    }
}
