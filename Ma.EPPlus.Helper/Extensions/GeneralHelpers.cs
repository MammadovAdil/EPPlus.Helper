using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Ma.EPPlus.Helper
{
    /// <summary>
    /// General helper methods.
    /// </summary>
    public static class GeneralHelpers
    {
        /// <summary>
        /// Rename headers of worksheet.
        /// </summary>
        /// <exception cref="ArgumentNullException">
        /// When worksheet or headerMap is null.
        /// </exception>
        /// <param name="worksheet">Worksheet to rename headers.</param>
        /// <param name="headerMap">
        /// Header map. Keys are current headers and values are desired headers.
        /// </param>
        public static void RenameHeaders(
            this ExcelWorksheet worksheet,
            Dictionary<string, string> headerMap)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (headerMap == null)
                throw new ArgumentNullException(nameof(headerMap));

            // Get header cells
            var headerCells = worksheet
                .Cells[1, 1, 1, worksheet.Dimension.End.Column];

            // Loop thorugh provided header map dictionary and rename headers.
            foreach (KeyValuePair<string, string> item in headerMap)
            {
                var cell = headerCells.FirstOrDefault(
                    m => m.Text.Equals(item.Key, StringComparison.InvariantCultureIgnoreCase));
                if (cell != null)
                    cell.Value = item.Value;
            }
        }

        /// <summary>
        /// Add headers to excel worksheet.
        /// </summary>
        /// <exception cref="ArgumentNullException">
        /// When worksheet or headers is null.
        /// </exception>
        /// <param name="worksheet">Excell worksheet to add headers.</param>
        /// <param name="headers">Headers to add to excel worksheet.</param>
        public static void AddHeaders(
            this ExcelWorksheet worksheet,
            List<string> headers)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            if (headers == null)
                throw new ArgumentNullException(nameof(headers));

            // Loop and add headers
            for (int i = 0; i < headers.Count; i++)
            {
                // Get heaer cell. Starts from 1 so it is i + 1
                var cell = worksheet.Cells[1, i + 1];
                if (cell != null)
                    cell.Value = headers[i];
            }
        }
    }
}
