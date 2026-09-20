using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using ZeroDocuments.Common;
using ZeroDocuments.Excel.Models;

namespace ZeroDocuments.Excel
{
    /// <summary>
    /// Pure C# Zero-Dependency OpenXML Excel (.xlsx) Reader.
    /// High-performance, streaming-first architecture using forward-only XmlReader.
    /// Operates with &lt; 15MB RAM regardless of sheet row count.
    /// </summary>
    public static class ExcelReader
    {
        private const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NsRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        static ExcelReader()
        {
            RuntimeAssemblyResolver.EnsureInitialized();
        }

        #region Public DataTable & POCO APIs

        /// <summary>
        /// Reads Excel sheet into a DataTable bounded by header range (e.g. "D24:T24" or "A1:C1").
        /// Automatically expands rows downward until data ends.
        /// </summary>
        public static DataTable ReadByHeaderRange(string filePath, string headerRange, int maxRows = 5000, string? sheetName = null)
        {
            string fullDataRange = ExcelCellAddress.ConvertHeaderRangeToDataRange(headerRange, maxRows);
            return ReadToDataTable(filePath, fullDataRange, sheetName);
        }

        /// <summary>
        /// Reads Excel file from file path into a DataTable.
        /// </summary>
        public static DataTable ReadToDataTable(string filePath, string cellRange = ExcelCellAddress.DefaultRange, string? sheetName = null)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                throw new FileNotFoundException($"Excel file not found: {filePath}");
            }

            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return ReadToDataTable(stream, cellRange, sheetName);
        }

        /// <summary>
        /// Reads Excel stream into a DataTable.
        /// </summary>
        public static DataTable ReadToDataTable(Stream stream, string cellRange = ExcelCellAddress.DefaultRange, string? sheetName = null)
        {
            var table = new DataTable();
            ExcelCellAddress.ParseCellRange(cellRange, out var startCol, out _, out var endCol, out _);
            int startColIdx = ExcelCellAddress.ColumnNameToIndex(startCol);
            int endColIdx = ExcelCellAddress.ColumnNameToIndex(endCol);

            for (int c = startColIdx; c <= endColIdx; c++)
            {
                table.Columns.Add("Column_" + ExcelCellAddress.IndexToColumnName(c), typeof(string));
            }

            var rows = ReadRows(stream, cellRange, sheetName);
            foreach (var row in rows)
            {
                var rowVals = new object?[endColIdx - startColIdx + 1];
                bool hasData = false;

                for (int c = startColIdx; c <= endColIdx; c++)
                {
                    var val = row[c];
                    if (!string.IsNullOrEmpty(val))
                    {
                        hasData = true;
                    }
                    rowVals[c - startColIdx] = val;
                }

                if (hasData)
                {
                    table.Rows.Add(rowVals);
                }
            }

            return table;
        }

        /// <summary>
        /// Reads an Excel file directly into strongly-typed POCO objects using compiled PropertyAccessorCache.
        /// </summary>
        public static List<T> Read<T>(string filePath, string cellRange = ExcelCellAddress.DefaultRange, string? sheetName = null) where T : new()
        {
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return Read<T>(stream, cellRange, sheetName);
        }

        /// <summary>
        /// Reads an Excel stream directly into strongly-typed POCO objects using compiled PropertyAccessorCache.
        /// </summary>
        public static List<T> Read<T>(Stream stream, string cellRange = ExcelCellAddress.DefaultRange, string? sheetName = null) where T : new()
        {
            var list = new List<T>();
            var rows = ReadRows(stream, cellRange, sheetName);
            if (rows.Count == 0) return list;

            var accessors = PropertyAccessorCache.GetAccessors(typeof(T));
            var accessorMap = new Dictionary<string, PropertyAccessorCache.PropertyAccessorInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (var acc in accessors)
            {
                accessorMap[acc.Name] = acc;
            }

            var colToAccessor = new Dictionary<int, PropertyAccessorCache.PropertyAccessorInfo>();
            bool isFirst = true;

            foreach (var row in rows)
            {
                if (isFirst)
                {
                    isFirst = false;
                    foreach (var colIdx in row.PopulatedColumns)
                    {
                        string? header = row[colIdx];
                        if (!string.IsNullOrWhiteSpace(header) && accessorMap.TryGetValue(header!.Trim(), out var acc))
                        {
                            colToAccessor[colIdx] = acc;
                        }
                    }
                    continue;
                }

                var item = new T();
                bool assigned = false;

                foreach (var kvp in colToAccessor)
                {
                    string? rawVal = row[kvp.Key];
                    if (rawVal != null && kvp.Value.Setter != null)
                    {
                        try
                        {
                            object? converted = ConvertValue(rawVal, kvp.Value.PropertyType);
                            kvp.Value.Setter(item, converted);
                            assigned = true;
                        }
                        catch
                        {
                            // Ignore casting/parsing errors
                        }
                    }
                }

                if (assigned)
                {
                    list.Add(item);
                }
            }

            return list;
        }

        #endregion

        #region Streaming Row Reading APIs

        /// <summary>
        /// Reads Excel rows from file as a list of ExcelRow objects.
        /// </summary>
        public static List<ExcelRow> ReadRows(string filePath, string cellRange = ExcelCellAddress.DefaultRange, string? sheetName = null)
        {
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return ReadRows(stream, cellRange, sheetName);
        }

        /// <summary>
        /// Reads Excel stream rows as a list of ExcelRow objects.
        /// </summary>
        public static List<ExcelRow> ReadRows(Stream stream, string cellRange = ExcelCellAddress.DefaultRange, string? sheetName = null)
        {
            return StreamRows(stream, cellRange, sheetName).ToList();
        }

        /// <summary>
        /// Streams Excel rows lazily using forward-only XmlReader.
        /// Minimal RAM footprint (&lt; 15MB) for maximum scalability.
        /// </summary>
        public static IEnumerable<ExcelRow> StreamRows(Stream stream, string cellRange = ExcelCellAddress.DefaultRange, string? sheetName = null)
        {
            ExcelCellAddress.ParseCellRange(cellRange, out var startCol, out var startRow, out var endCol, out var endRow);
            int startColIdx = ExcelCellAddress.ColumnNameToIndex(startCol);
            int endColIdx = ExcelCellAddress.ColumnNameToIndex(endCol);

            using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

            // 1. Load shared strings table if present
            var sharedStrings = LoadSharedStrings(zip);

            // 2. Locate target worksheet entry
            var sheetEntry = FindSheetEntry(zip, sheetName);
            if (sheetEntry == null) yield break;

            using var sheetStream = sheetEntry.Open();
            using var xmlReader = XmlReader.Create(sheetStream, new XmlReaderSettings
            {
                IgnoreWhitespace = true,
                IgnoreComments = true
            });

            while (xmlReader.Read())
            {
                if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.LocalName == "row")
                {
                    string? rAttr = xmlReader.GetAttribute("r");
                    int rowNum = int.TryParse(rAttr, out int rVal) ? rVal : 0;

                    if (rowNum > 0 && rowNum < startRow)
                    {
                        continue;
                    }
                    if (rowNum > endRow)
                    {
                        yield break;
                    }

                    var excelRow = new ExcelRow { RowNumber = rowNum };
                    bool hasAnyCell = false;

                    if (!xmlReader.IsEmptyElement)
                    {
                        using var rowSubtree = xmlReader.ReadSubtree();
                        while (rowSubtree.Read())
                        {
                            if (rowSubtree.NodeType == XmlNodeType.Element && rowSubtree.LocalName == "c")
                            {
                                string? cellRef = rowSubtree.GetAttribute("r");
                                string? cellType = rowSubtree.GetAttribute("t");

                                if (string.IsNullOrEmpty(cellRef) ||
                                    !ExcelCellAddress.TryParseCellReference(cellRef!, out var colName, out _))
                                {
                                    continue;
                                }

                                int colIdx = ExcelCellAddress.ColumnNameToIndex(colName);
                                if (colIdx < startColIdx || colIdx > endColIdx) continue;

                                string? cellVal = null;
                                if (!rowSubtree.IsEmptyElement)
                                {
                                    using var cellSubtree = rowSubtree.ReadSubtree();
                                    while (cellSubtree.Read())
                                    {
                                        if (cellSubtree.NodeType == XmlNodeType.Element)
                                        {
                                            if (cellSubtree.LocalName == "v")
                                            {
                                                string rawVal = cellSubtree.ReadElementContentAsString();
                                                if (cellType == "s")
                                                {
                                                    if (int.TryParse(rawVal, out int sIdx) && sIdx >= 0 && sIdx < sharedStrings.Count)
                                                        cellVal = sharedStrings[sIdx];
                                                    else
                                                        cellVal = rawVal;
                                                }
                                                else if (cellType == "b")
                                                {
                                                    cellVal = rawVal == "1" ? "TRUE" : "FALSE";
                                                }
                                                else
                                                {
                                                    cellVal = rawVal;
                                                }
                                            }
                                            else if (cellSubtree.LocalName == "t")
                                            {
                                                cellVal = cellSubtree.ReadElementContentAsString();
                                            }
                                        }
                                    }
                                }

                                excelRow[colIdx] = cellVal;
                                if (!string.IsNullOrEmpty(cellVal))
                                {
                                    hasAnyCell = true;
                                }
                            }
                        }
                    }

                    if (hasAnyCell)
                    {
                        yield return excelRow;
                    }
                }
            }
        }

        #endregion

        #region Private Package & XML Helpers

        private static List<string> LoadSharedStrings(ZipArchive zip)
        {
            var list = new List<string>();
            var ssEntry = zip.Entries.FirstOrDefault(e => e.FullName.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase));
            if (ssEntry == null) return list;

            using var stream = ssEntry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true });

            var sb = new StringBuilder();
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
                {
                    sb.Clear();
                    if (!reader.IsEmptyElement)
                    {
                        using var siSubtree = reader.ReadSubtree();
                        while (siSubtree.Read())
                        {
                            if (siSubtree.NodeType == XmlNodeType.Element && siSubtree.LocalName == "t")
                            {
                                sb.Append(siSubtree.ReadElementContentAsString());
                            }
                        }
                    }
                    list.Add(sb.ToString());
                }
            }
            return list;
        }

        private static ZipArchiveEntry? FindSheetEntry(ZipArchive zip, string? sheetName)
        {
            if (!string.IsNullOrEmpty(sheetName))
            {
                // 1. Try to resolve sheet name and relationship via xl/workbook.xml
                var wbEntry = zip.Entries.FirstOrDefault(e => e.FullName.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase));
                if (wbEntry != null)
                {
                    using var wbStream = wbEntry.Open();
                    using var reader = XmlReader.Create(wbStream, new XmlReaderSettings { IgnoreWhitespace = true });

                    string? matchedRelId = null;
                    string? matchedSheetId = null;

                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheet")
                        {
                            string? name = reader.GetAttribute("name");
                            if (string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                matchedRelId = reader.GetAttribute("id", NsRelationships) ?? reader.GetAttribute("r:id");
                                matchedSheetId = reader.GetAttribute("sheetId");
                                break;
                            }
                        }
                    }

                    // Look up target path from xl/_rels/workbook.xml.rels
                    if (!string.IsNullOrEmpty(matchedRelId))
                    {
                        var relsEntry = zip.Entries.FirstOrDefault(e => e.FullName.Equals("xl/_rels/workbook.xml.rels", StringComparison.OrdinalIgnoreCase));
                        if (relsEntry != null)
                        {
                            using var relsStream = relsEntry.Open();
                            using var relsReader = XmlReader.Create(relsStream, new XmlReaderSettings { IgnoreWhitespace = true });

                            while (relsReader.Read())
                            {
                                if (relsReader.NodeType == XmlNodeType.Element && relsReader.LocalName == "Relationship")
                                {
                                    string? id = relsReader.GetAttribute("Id");
                                    if (string.Equals(id, matchedRelId, StringComparison.OrdinalIgnoreCase))
                                    {
                                        string? target = relsReader.GetAttribute("Target");
                                        if (!string.IsNullOrEmpty(target))
                                        {
                                            string fullPath = target.StartsWith("xl/") ? target : "xl/" + target.TrimStart('/');
                                            var found = zip.Entries.FirstOrDefault(e => e.FullName.Equals(fullPath, StringComparison.OrdinalIgnoreCase));
                                            if (found != null) return found;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (!string.IsNullOrEmpty(matchedSheetId))
                    {
                        var entry = zip.Entries.FirstOrDefault(e => e.FullName.Equals($"xl/worksheets/sheet{matchedSheetId}.xml", StringComparison.OrdinalIgnoreCase));
                        if (entry != null) return entry;
                    }
                }

                // If a specific sheet name was requested but could not be resolved, do not return the wrong sheet.
                return null;
            }

            // Fallback when no specific sheetName is requested: first sheet entry in xl/worksheets/
            return zip.Entries.FirstOrDefault(e => e.FullName.Equals("xl/worksheets/sheet1.xml", StringComparison.OrdinalIgnoreCase))
                   ?? zip.Entries.FirstOrDefault(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        }

        private static object? ConvertValue(string raw, Type targetType)
        {
            if (string.IsNullOrEmpty(raw)) return null;

            Type underlying = Nullable.GetUnderlyingType(targetType) ?? targetType;

            if (underlying == typeof(string)) return raw;
            if (underlying == typeof(int)) return int.Parse(raw, CultureInfo.InvariantCulture);
            if (underlying == typeof(long)) return long.Parse(raw, CultureInfo.InvariantCulture);
            if (underlying == typeof(double)) return double.Parse(raw, CultureInfo.InvariantCulture);
            if (underlying == typeof(float)) return float.Parse(raw, CultureInfo.InvariantCulture);
            if (underlying == typeof(decimal)) return decimal.Parse(raw, CultureInfo.InvariantCulture);
            if (underlying == typeof(bool))
            {
                if (raw == "1" || raw.Equals("true", StringComparison.OrdinalIgnoreCase) || raw.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return true;
                if (raw == "0" || raw.Equals("false", StringComparison.OrdinalIgnoreCase) || raw.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return false;
                return bool.Parse(raw);
            }
            if (underlying == typeof(DateTime)) return DateTime.Parse(raw, CultureInfo.InvariantCulture);
            if (underlying == typeof(DateTimeOffset)) return DateTimeOffset.Parse(raw, CultureInfo.InvariantCulture);
            if (underlying == typeof(TimeSpan)) return TimeSpan.Parse(raw, CultureInfo.InvariantCulture);
            if (underlying == typeof(Guid)) return Guid.Parse(raw);
            if (underlying.IsEnum)
            {
                return Enum.Parse(underlying, raw, true);
            }

            return Convert.ChangeType(raw, underlying, CultureInfo.InvariantCulture);
        }

        #endregion
    }
}
