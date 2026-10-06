using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using ZeroDocuments.Common;
using ZeroDocuments.Excel.Internal;

namespace ZeroDocuments.Excel
{
    /// <summary>
    /// Pure C# Zero-Dependency OpenXML Excel (.xlsx) Writer.
    /// Operates without external DLLs (No EPPlus, ClosedXML, or DocumentFormat.OpenXml required).
    /// Rows are streamed directly into the package: sources are enumerated exactly once and never buffered.
    /// </summary>
    public static class ExcelWriter
    {
        /// <summary>
        /// Gets or sets whether formula injection protection (CWE-1236) is enabled.
        /// When true, strings beginning with =, +, -, @, \t, or \r are prefixed with a single quote.
        /// Default is true.
        /// </summary>
        /// <remarks>
        /// This is process-wide state. Prefer <see cref="ExcelWorkbookBuilder.FormulaInjectionProtection"/>
        /// (per-instance) when different call sites require different policies.
        /// </remarks>
        public static bool FormulaInjectionProtection { get; set; } = true;

        static ExcelWriter()
        {
            RuntimeAssemblyResolver.EnsureInitialized();
        }

        #region Public Write APIs

        /// <summary>
        /// Writes a DataTable to an Excel (.xlsx) file on disk.
        /// </summary>
        public static void WriteToFile(string filePath, DataTable table, string sheetName = "Sheet1", bool includeHeaders = true)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            using var stream = CreateFile(filePath);
            WriteToStream(stream, table, sheetName, includeHeaders);
        }

        /// <summary>
        /// Writes a DataTable to a stream in OpenXML Excel (.xlsx) format.
        /// </summary>
        public static void WriteToStream(Stream stream, DataTable table, string sheetName = "Sheet1", bool includeHeaders = true)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (table == null) throw new ArgumentNullException(nameof(table));

            WriteRowsToStream(stream, TabularSource.FromDataTable(table), includeHeaders ? TabularSource.GetHeaders(table) : null, sheetName);
        }

        /// <summary>
        /// Writes a collection of objects to an Excel (.xlsx) file on disk.
        /// Public properties are mapped to columns.
        /// </summary>
        public static void WriteToFile<T>(string filePath, IEnumerable<T> data, string sheetName = "Sheet1", bool includeHeaders = true)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            using var stream = CreateFile(filePath);
            WriteToStream(stream, data, sheetName, includeHeaders);
        }

        /// <summary>
        /// Writes a collection of objects to a stream in OpenXML Excel (.xlsx) format.
        /// </summary>
        public static void WriteToStream<T>(Stream stream, IEnumerable<T> data, string sheetName = "Sheet1", bool includeHeaders = true)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (data == null) throw new ArgumentNullException(nameof(data));

            var accessors = PropertyAccessorCache.GetAccessors(typeof(T));
            WriteRowsToStream(stream, TabularSource.FromObjects(data, accessors), includeHeaders ? TabularSource.GetHeaders(accessors) : null, sheetName);
        }

        /// <summary>
        /// Writes raw 2D grid rows to an Excel (.xlsx) file on disk.
        /// </summary>
        public static void WriteToFile(string filePath, IEnumerable<IReadOnlyList<object?>> rows, IReadOnlyList<string>? headers = null, string sheetName = "Sheet1")
        {
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            using var stream = CreateFile(filePath);
            WriteRowsToStream(stream, rows, headers, sheetName);
        }

        /// <summary>
        /// Writes raw 2D grid rows to a stream in OpenXML Excel (.xlsx) format.
        /// </summary>
        public static void WriteRowsToStream(Stream stream, IEnumerable<IReadOnlyList<object?>> rows, IReadOnlyList<string>? headers = null, string sheetName = "Sheet1")
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (rows == null) throw new ArgumentNullException(nameof(rows));

            var sheet = new WorksheetSpec
            {
                Name = SpreadsheetXml.SanitizeSheetName(sheetName, "Sheet1"),
                Headers = headers,
                Rows = rows,
                HeaderBold = false
            };

            XlsxPackageWriter.Write(stream, new[] { sheet }, FormulaInjectionProtection);
        }

        #endregion

        #region Helpers

        /// <summary>
        /// Sanitizes text to prevent formula injection attacks (CWE-1236).
        /// Prefixes a single quote if the first character is =, +, -, @, \t, or \r.
        /// </summary>
        public static string SanitizeFormulaInjection(string? text) => FormulaInjectionGuard.Sanitize(text);

        internal static FileStream CreateFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));

            string? directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            return new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
        }

        #endregion
    }
}
