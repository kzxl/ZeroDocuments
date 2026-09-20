using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using ZeroDocuments.Common;
using ZeroDocuments.Excel.Models;

namespace ZeroDocuments.Excel
{
    /// <summary>
    /// Fluent multi-sheet OpenXML Excel (.xlsx) workbook builder.
    /// Pure C# BCL implementation with zero external dependencies.
    /// </summary>
    public sealed class ExcelWorkbookBuilder : IDisposable
    {
        private const string NsSpreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NsRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private readonly List<WorksheetDefinition> _sheets = new List<WorksheetDefinition>();

        /// <summary>
        /// Gets or sets whether formula injection protection (CWE-1236) is enabled.
        /// Default is true.
        /// </summary>
        public bool FormulaInjectionProtection { get; set; } = true;

        private sealed class WorksheetDefinition
        {
            public string Name { get; set; } = string.Empty;
            public IReadOnlyList<string>? Headers { get; set; }
            public List<IReadOnlyList<object?>> Rows { get; } = new List<IReadOnlyList<object?>>();
            public bool HeaderBold { get; set; } = true;
        }

        static ExcelWorkbookBuilder()
        {
            RuntimeAssemblyResolver.EnsureInitialized();
        }

        #region Fluent AddSheet APIs

        /// <summary>
        /// Adds a new worksheet populated from a System.Data.DataTable.
        /// </summary>
        public ExcelWorkbookBuilder AddSheet(string sheetName, DataTable table, bool includeHeaders = true, bool headerBold = true)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));

            var headers = new List<string>(table.Columns.Count);
            foreach (DataColumn col in table.Columns)
            {
                headers.Add(col.ColumnName);
            }

            var sheet = new WorksheetDefinition
            {
                Name = SanitizeSheetName(sheetName),
                Headers = includeHeaders ? headers : null,
                HeaderBold = headerBold
            };

            foreach (DataRow row in table.Rows)
            {
                var values = new object?[table.Columns.Count];
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    values[i] = row[i] == DBNull.Value ? null : row[i];
                }
                sheet.Rows.Add(values);
            }

            _sheets.Add(sheet);
            return this;
        }

        /// <summary>
        /// Adds a new worksheet populated from an IEnumerable of strongly-typed POCO objects.
        /// </summary>
        public ExcelWorkbookBuilder AddSheet<T>(string sheetName, IEnumerable<T> data, bool includeHeaders = true, bool headerBold = true)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));

            var accessors = PropertyAccessorCache.GetAccessors(typeof(T));
            var headers = new List<string>(accessors.Length);
            foreach (var acc in accessors)
            {
                headers.Add(acc.Name);
            }

            var sheet = new WorksheetDefinition
            {
                Name = SanitizeSheetName(sheetName),
                Headers = includeHeaders ? headers : null,
                HeaderBold = headerBold
            };

            foreach (var item in data)
            {
                if (item == null) continue;
                var values = new object?[accessors.Length];
                for (int i = 0; i < accessors.Length; i++)
                {
                    values[i] = accessors[i].Getter(item);
                }
                sheet.Rows.Add(values);
            }

            _sheets.Add(sheet);
            return this;
        }

        /// <summary>
        /// Adds a new worksheet populated from raw 2D grid rows.
        /// </summary>
        public ExcelWorkbookBuilder AddSheet(string sheetName, IEnumerable<IReadOnlyList<object?>> rows, IReadOnlyList<string>? headers = null, bool headerBold = true)
        {
            if (rows == null) throw new ArgumentNullException(nameof(rows));

            var sheet = new WorksheetDefinition
            {
                Name = SanitizeSheetName(sheetName),
                Headers = headers,
                HeaderBold = headerBold
            };

            foreach (var r in rows)
            {
                sheet.Rows.Add(r);
            }

            _sheets.Add(sheet);
            return this;
        }

        #endregion

        #region Save & Export APIs

        /// <summary>
        /// Saves the workbook to an Excel (.xlsx) file on disk.
        /// </summary>
        public void Save(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));

            string? dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
            Save(stream);
        }

        /// <summary>
        /// Saves the workbook to a destination stream in OpenXML Excel (.xlsx) format.
        /// </summary>
        public void Save(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));

            if (_sheets.Count == 0)
            {
                // Ensure at least 1 worksheet exists
                AddSheet("Sheet1", Array.Empty<IReadOnlyList<object?>>());
            }

            using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);

            // 1. [Content_Types].xml
            CreateContentTypesEntry(zip, _sheets.Count);

            // 2. _rels/.rels
            CreateGlobalRelsEntry(zip);

            // 3. xl/workbook.xml
            CreateWorkbookEntry(zip, _sheets);

            // 4. xl/_rels/workbook.xml.rels
            CreateWorkbookRelsEntry(zip, _sheets.Count);

            // 5. xl/styles.xml
            CreateStylesEntry(zip);

            // 6. xl/worksheets/sheet{N}.xml
            for (int i = 0; i < _sheets.Count; i++)
            {
                CreateWorksheetEntry(zip, i + 1, _sheets[i]);
            }
        }

        /// <summary>
        /// Returns the entire workbook package as a byte array.
        /// </summary>
        public byte[] ToArray()
        {
            using var ms = new MemoryStream();
            Save(ms);
            return ms.ToArray();
        }

        public void Dispose()
        {
            _sheets.Clear();
        }

        #endregion

        #region OPC Package Generation

        private static void CreateContentTypesEntry(ZipArchive zip, int sheetCount)
        {
            var entry = zip.CreateEntry("[Content_Types].xml", CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = CreateXmlWriter(stream);

            writer.WriteStartDocument(true);
            writer.WriteStartElement("Types", "http://schemas.openxmlformats.org/package/2006/content-types");

            writer.WriteStartElement("Default");
            writer.WriteAttributeString("Extension", "rels");
            writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
            writer.WriteEndElement();

            writer.WriteStartElement("Default");
            writer.WriteAttributeString("Extension", "xml");
            writer.WriteAttributeString("ContentType", "application/xml");
            writer.WriteEndElement();

            writer.WriteStartElement("Override");
            writer.WriteAttributeString("PartName", "/xl/workbook.xml");
            writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            writer.WriteEndElement();

            writer.WriteStartElement("Override");
            writer.WriteAttributeString("PartName", "/xl/styles.xml");
            writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            writer.WriteEndElement();

            for (int i = 1; i <= sheetCount; i++)
            {
                writer.WriteStartElement("Override");
                writer.WriteAttributeString("PartName", $"/xl/worksheets/sheet{i}.xml");
                writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                writer.WriteEndElement();
            }

            writer.WriteEndElement(); // Types
            writer.WriteEndDocument();
        }

        private static void CreateGlobalRelsEntry(ZipArchive zip)
        {
            var entry = zip.CreateEntry("_rels/.rels", CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = CreateXmlWriter(stream);

            writer.WriteStartDocument(true);
            writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");

            writer.WriteStartElement("Relationship");
            writer.WriteAttributeString("Id", "rId1");
            writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
            writer.WriteAttributeString("Target", "xl/workbook.xml");
            writer.WriteEndElement();

            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        private static void CreateWorkbookEntry(ZipArchive zip, List<WorksheetDefinition> sheets)
        {
            var entry = zip.CreateEntry("xl/workbook.xml", CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = CreateXmlWriter(stream);

            writer.WriteStartDocument(true);
            writer.WriteStartElement("workbook", NsSpreadsheet);
            writer.WriteAttributeString("xmlns", "r", null, NsRelationships);

            writer.WriteStartElement("sheets");
            for (int i = 0; i < sheets.Count; i++)
            {
                writer.WriteStartElement("sheet");
                writer.WriteAttributeString("name", sheets[i].Name);
                writer.WriteAttributeString("sheetId", (i + 1).ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("id", NsRelationships, $"rId{i + 1}");
                writer.WriteEndElement();
            }
            writer.WriteEndElement(); // sheets

            writer.WriteEndElement(); // workbook
            writer.WriteEndDocument();
        }

        private static void CreateWorkbookRelsEntry(ZipArchive zip, int sheetCount)
        {
            var entry = zip.CreateEntry("xl/_rels/workbook.xml.rels", CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = CreateXmlWriter(stream);

            writer.WriteStartDocument(true);
            writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");

            for (int i = 1; i <= sheetCount; i++)
            {
                writer.WriteStartElement("Relationship");
                writer.WriteAttributeString("Id", $"rId{i}");
                writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                writer.WriteAttributeString("Target", $"worksheets/sheet{i}.xml");
                writer.WriteEndElement();
            }

            // Styles relationship
            writer.WriteStartElement("Relationship");
            writer.WriteAttributeString("Id", $"rId{sheetCount + 1}");
            writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            writer.WriteAttributeString("Target", "styles.xml");
            writer.WriteEndElement();

            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        private static void CreateStylesEntry(ZipArchive zip)
        {
            var entry = zip.CreateEntry("xl/styles.xml", CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = CreateXmlWriter(stream);

            writer.WriteStartDocument(true);
            writer.WriteStartElement("styleSheet", NsSpreadsheet);

            // fonts
            writer.WriteStartElement("fonts");
            writer.WriteAttributeString("count", "2");

            // font 0: Regular
            writer.WriteStartElement("font");
            writer.WriteStartElement("sz");
            writer.WriteAttributeString("val", "11");
            writer.WriteEndElement();
            writer.WriteStartElement("name");
            writer.WriteAttributeString("val", "Calibri");
            writer.WriteEndElement();
            writer.WriteEndElement();

            // font 1: Bold
            writer.WriteStartElement("font");
            writer.WriteStartElement("b");
            writer.WriteEndElement();
            writer.WriteStartElement("sz");
            writer.WriteAttributeString("val", "11");
            writer.WriteEndElement();
            writer.WriteStartElement("name");
            writer.WriteAttributeString("val", "Calibri");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteEndElement(); // fonts

            // fills
            writer.WriteStartElement("fills");
            writer.WriteAttributeString("count", "2");
            writer.WriteStartElement("fill");
            writer.WriteStartElement("patternFill");
            writer.WriteAttributeString("patternType", "none");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteStartElement("fill");
            writer.WriteStartElement("patternFill");
            writer.WriteAttributeString("patternType", "gray125");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement(); // fills

            // borders
            writer.WriteStartElement("borders");
            writer.WriteAttributeString("count", "1");
            writer.WriteStartElement("border");
            writer.WriteElementString("left", NsSpreadsheet, "");
            writer.WriteElementString("right", NsSpreadsheet, "");
            writer.WriteElementString("top", NsSpreadsheet, "");
            writer.WriteElementString("bottom", NsSpreadsheet, "");
            writer.WriteElementString("diagonal", NsSpreadsheet, "");
            writer.WriteEndElement();
            writer.WriteEndElement(); // borders

            // cellStyleXfs
            writer.WriteStartElement("cellStyleXfs");
            writer.WriteAttributeString("count", "1");
            writer.WriteStartElement("xf");
            writer.WriteAttributeString("numFmtId", "0");
            writer.WriteAttributeString("fontId", "0");
            writer.WriteAttributeString("fillId", "0");
            writer.WriteAttributeString("borderId", "0");
            writer.WriteEndElement();
            writer.WriteEndElement();

            // cellXfs
            writer.WriteStartElement("cellXfs");
            writer.WriteAttributeString("count", "2");

            // xf 0: Default
            writer.WriteStartElement("xf");
            writer.WriteAttributeString("numFmtId", "0");
            writer.WriteAttributeString("fontId", "0");
            writer.WriteAttributeString("fillId", "0");
            writer.WriteAttributeString("borderId", "0");
            writer.WriteAttributeString("xfId", "0");
            writer.WriteEndElement();

            // xf 1: Header Bold
            writer.WriteStartElement("xf");
            writer.WriteAttributeString("numFmtId", "0");
            writer.WriteAttributeString("fontId", "1");
            writer.WriteAttributeString("fillId", "0");
            writer.WriteAttributeString("borderId", "0");
            writer.WriteAttributeString("xfId", "0");
            writer.WriteAttributeString("applyFont", "1");
            writer.WriteEndElement();

            writer.WriteEndElement(); // cellXfs

            writer.WriteEndElement(); // styleSheet
            writer.WriteEndDocument();
        }

        private void CreateWorksheetEntry(ZipArchive zip, int sheetIndex, WorksheetDefinition sheet)
        {
            var entry = zip.CreateEntry($"xl/worksheets/sheet{sheetIndex}.xml", CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = CreateXmlWriter(stream);

            writer.WriteStartDocument(true);
            writer.WriteStartElement("worksheet", NsSpreadsheet);

            writer.WriteStartElement("sheetData");

            int currentRowIndex = 1;

            // 1. Write Header Row if provided
            if (sheet.Headers != null && sheet.Headers.Count > 0)
            {
                writer.WriteStartElement("row");
                writer.WriteAttributeString("r", currentRowIndex.ToString(CultureInfo.InvariantCulture));

                for (int col = 0; col < sheet.Headers.Count; col++)
                {
                    string colLetter = ExcelCellAddress.IndexToColumnName(col + 1);
                    string cellRef = $"{colLetter}{currentRowIndex}";
                    WriteCellString(writer, cellRef, sheet.Headers[col], styleIndex: sheet.HeaderBold ? 1 : 0);
                }

                writer.WriteEndElement(); // row
                currentRowIndex++;
            }

            // 2. Write Data Rows
            foreach (var rowValues in sheet.Rows)
            {
                if (rowValues == null)
                {
                    currentRowIndex++;
                    continue;
                }

                writer.WriteStartElement("row");
                writer.WriteAttributeString("r", currentRowIndex.ToString(CultureInfo.InvariantCulture));

                for (int col = 0; col < rowValues.Count; col++)
                {
                    var val = rowValues[col];
                    if (val == null) continue;

                    string colLetter = ExcelCellAddress.IndexToColumnName(col + 1);
                    string cellRef = $"{colLetter}{currentRowIndex}";

                    WriteCellValue(writer, cellRef, val);
                }

                writer.WriteEndElement(); // row
                currentRowIndex++;
            }

            writer.WriteEndElement(); // sheetData
            writer.WriteEndElement(); // worksheet
            writer.WriteEndDocument();
        }

        private void WriteCellValue(XmlWriter writer, string cellRef, object val)
        {
            switch (val)
            {
                case bool b:
                    writer.WriteStartElement("c");
                    writer.WriteAttributeString("r", cellRef);
                    writer.WriteAttributeString("t", "b");
                    writer.WriteElementString("v", NsSpreadsheet, b ? "1" : "0");
                    writer.WriteEndElement();
                    break;

                case sbyte or byte or short or ushort or int or uint or long or ulong:
                    writer.WriteStartElement("c");
                    writer.WriteAttributeString("r", cellRef);
                    writer.WriteElementString("v", NsSpreadsheet, Convert.ToString(val, CultureInfo.InvariantCulture));
                    writer.WriteEndElement();
                    break;

                case float f:
                    writer.WriteStartElement("c");
                    writer.WriteAttributeString("r", cellRef);
                    writer.WriteElementString("v", NsSpreadsheet, f.ToString("R", CultureInfo.InvariantCulture));
                    writer.WriteEndElement();
                    break;

                case double d:
                    writer.WriteStartElement("c");
                    writer.WriteAttributeString("r", cellRef);
                    writer.WriteElementString("v", NsSpreadsheet, d.ToString("R", CultureInfo.InvariantCulture));
                    writer.WriteEndElement();
                    break;

                case decimal dec:
                    writer.WriteStartElement("c");
                    writer.WriteAttributeString("r", cellRef);
                    writer.WriteElementString("v", NsSpreadsheet, dec.ToString(CultureInfo.InvariantCulture));
                    writer.WriteEndElement();
                    break;

                case DateTime dt:
                    string dateStr = dt.TimeOfDay == TimeSpan.Zero
                        ? dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                        : dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    WriteCellString(writer, cellRef, dateStr);
                    break;

                case DateTimeOffset dto:
                    WriteCellString(writer, cellRef, dto.ToString("yyyy-MM-dd HH:mm:ss zzz", CultureInfo.InvariantCulture));
                    break;

                default:
                    WriteCellString(writer, cellRef, val.ToString() ?? string.Empty);
                    break;
            }
        }

        private void WriteCellString(XmlWriter writer, string cellRef, string rawText, int styleIndex = 0)
        {
            string safeText = FormulaInjectionProtection ? SanitizeFormulaInjection(rawText) : rawText;
            string cleanText = SanitizeXmlText(safeText);

            writer.WriteStartElement("c");
            writer.WriteAttributeString("r", cellRef);
            writer.WriteAttributeString("t", "inlineStr");
            if (styleIndex > 0)
            {
                writer.WriteAttributeString("s", styleIndex.ToString(CultureInfo.InvariantCulture));
            }

            writer.WriteStartElement("is");
            writer.WriteStartElement("t");
            if (cleanText.StartsWith(" ") || cleanText.EndsWith(" "))
            {
                writer.WriteAttributeString("xml", "space", null, "preserve");
            }
            writer.WriteString(cleanText);
            writer.WriteEndElement(); // t
            writer.WriteEndElement(); // is

            writer.WriteEndElement(); // c
        }

        private static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Sheet";

            var sb = new StringBuilder(name.Length);
            foreach (char ch in name)
            {
                if (ch == '\\' || ch == '/' || ch == '?' || ch == '*' || ch == ':' || ch == '[' || ch == ']')
                    sb.Append('_');
                else
                    sb.Append(ch);
            }

            string result = sb.ToString().Trim();
            if (result.Length > 31)
            {
                result = result.Substring(0, 31);
            }

            return string.IsNullOrEmpty(result) ? "Sheet" : result;
        }

        private static string SanitizeFormulaInjection(string? text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            char firstChar = text![0];
            if (firstChar == '=' || firstChar == '+' || firstChar == '-' || firstChar == '@' || firstChar == '\t' || firstChar == '\r')
            {
                return "'" + text;
            }
            return text;
        }

        private static string SanitizeXmlText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            var sb = new StringBuilder(text.Length);
            foreach (char ch in text)
            {
                if (ch == 0x9 || ch == 0xA || ch == 0xD ||
                    (ch >= 0x20 && ch <= 0xD7FF) ||
                    (ch >= 0xE000 && ch <= 0xFFFD))
                {
                    sb.Append(ch);
                }
            }

            return sb.ToString();
        }

        private static XmlWriter CreateXmlWriter(Stream stream)
        {
            var settings = new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                OmitXmlDeclaration = false,
                Indent = false,
                CloseOutput = false
            };
            return XmlWriter.Create(stream, settings);
        }

        #endregion
    }
}
