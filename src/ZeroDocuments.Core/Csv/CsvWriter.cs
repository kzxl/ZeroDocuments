using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using ZeroDocuments.Common;

namespace ZeroDocuments.Csv
{
    /// <summary>
    /// Pure C# Zero-Dependency RFC 4180 compliant CSV Writer.
    /// Supports streaming, proper character escaping, and DataTable/Collection exports.
    /// Rows are written as they are enumerated; sources are never buffered in memory.
    /// </summary>
    public static class CsvWriter
    {
        /// <summary>
        /// Gets or sets whether formula injection protection (CWE-1236) is enabled.
        /// When true, text fields beginning with =, +, -, @, \t, or \r are prefixed with a single quote.
        /// Typed numeric, boolean, date and time values are never prefixed (e.g. -5 stays "-5").
        /// Default is true.
        /// </summary>
        public static bool FormulaInjectionProtection { get; set; } = true;

        /// <summary>
        /// Writes a DataTable to a CSV file.
        /// </summary>
        public static void WriteToFile(string filePath, DataTable table, char delimiter = ',', bool includeHeaders = true, Encoding? encoding = null)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            using var stream = CreateFile(filePath);
            WriteToStream(stream, table, delimiter, includeHeaders, encoding);
        }

        /// <summary>
        /// Writes a DataTable to a stream in CSV format.
        /// </summary>
        public static void WriteToStream(Stream stream, DataTable table, char delimiter = ',', bool includeHeaders = true, Encoding? encoding = null)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (table == null) throw new ArgumentNullException(nameof(table));

            WriteRowsToStream(stream, TabularSource.FromDataTable(table), includeHeaders ? TabularSource.GetHeaders(table) : null, delimiter, encoding);
        }

        /// <summary>
        /// Writes a collection of objects to a CSV file.
        /// </summary>
        public static void WriteToFile<T>(string filePath, IEnumerable<T> data, char delimiter = ',', bool includeHeaders = true, Encoding? encoding = null)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            using var stream = CreateFile(filePath);
            WriteToStream(stream, data, delimiter, includeHeaders, encoding);
        }

        /// <summary>
        /// Writes a collection of objects to a stream in CSV format.
        /// </summary>
        public static void WriteToStream<T>(Stream stream, IEnumerable<T> data, char delimiter = ',', bool includeHeaders = true, Encoding? encoding = null)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (data == null) throw new ArgumentNullException(nameof(data));

            var accessors = PropertyAccessorCache.GetAccessors(typeof(T));
            WriteRowsToStream(stream, TabularSource.FromObjects(data, accessors), includeHeaders ? TabularSource.GetHeaders(accessors) : null, delimiter, encoding);
        }

        /// <summary>
        /// Writes raw 2D grid rows to a stream in CSV format.
        /// </summary>
        public static void WriteRowsToStream(Stream stream, IEnumerable<IReadOnlyList<object?>> rows, IReadOnlyList<string>? headers = null, char delimiter = ',', Encoding? encoding = null)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            if (delimiter == '"' || delimiter == '\r' || delimiter == '\n')
                throw new ArgumentException("Delimiter cannot be a double quote, CR or LF.", nameof(delimiter));

            bool guard = FormulaInjectionProtection;
            using var writer = new StreamWriter(stream, encoding ?? new UTF8Encoding(true), 4096, leaveOpen: true);

            // 1. Header
            if (headers != null && headers.Count > 0)
            {
                for (int i = 0; i < headers.Count; i++)
                {
                    if (i > 0) writer.Write(delimiter);
                    writer.Write(EscapeCsvField(headers[i], delimiter, guard));
                }
                writer.WriteLine();
            }

            // 2. Rows
            foreach (var row in rows)
            {
                if (row == null)
                {
                    writer.WriteLine();
                    continue;
                }

                for (int i = 0; i < row.Count; i++)
                {
                    if (i > 0) writer.Write(delimiter);
                    writer.Write(FormatValue(row[i], delimiter, guard));
                }
                writer.WriteLine();
            }

            writer.Flush();
        }

        private static string FormatValue(object? val, char delimiter, bool guard)
        {
            switch (val)
            {
                case null:
                    return string.Empty;
                case DBNull _:
                    return string.Empty;
                case string s:
                    return EscapeCsvField(s, delimiter, guard);
                case DateTime dt:
                    // InvariantCulture: ':' in custom formats is the culture's time separator otherwise.
                    return EscapeCsvField(dt.TimeOfDay == TimeSpan.Zero
                        ? dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                        : dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), delimiter, sanitize: false);
                case sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal or bool or TimeSpan or DateTimeOffset:
                    // Typed values are data, not attacker-controlled text: never prefix (keeps negative numbers intact).
                    return EscapeCsvField(val.ToString(), delimiter, sanitize: false);
                default:
                    return EscapeCsvField(val.ToString(), delimiter, guard);
            }
        }

        private static string EscapeCsvField(string? field, char delimiter, bool sanitize)
        {
            if (string.IsNullOrEmpty(field)) return string.Empty;

            string processed = sanitize ? FormulaInjectionGuard.Sanitize(field) : field!;

            bool mustQuote = processed.IndexOf(delimiter) >= 0 ||
                             processed.IndexOf('"') >= 0 ||
                             processed.IndexOf('\n') >= 0 ||
                             processed.IndexOf('\r') >= 0;

            if (!mustQuote) return processed;

            // Double up internal quotes: " -> ""
            return "\"" + processed.Replace("\"", "\"\"") + "\"";
        }

        private static FileStream CreateFile(string filePath)
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
    }
}
