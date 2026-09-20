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
    /// Supports rich formatting, embedded DrawingML images, and conditional formatting.
    /// </summary>
    public sealed class ExcelWorkbookBuilder : IDisposable
    {
        private const string NsSpreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NsRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string NsDrawingSheet = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        private const string NsDrawingMain = "http://schemas.openxmlformats.org/drawingml/2006/main";

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
            public string? AutoFilterRange { get; set; }
            public bool AutoFilterEnabled { get; set; }
            public List<ExcelImage> Images { get; } = new List<ExcelImage>();
            public List<ExcelConditionalFormatRule> ConditionalFormatting { get; } = new List<ExcelConditionalFormatRule>();
        }

        private sealed class DxfStyleDefinition : IEquatable<DxfStyleDefinition>
        {
            public string? FillColorHex { get; set; }
            public string? FontColorHex { get; set; }
            public bool Bold { get; set; }

            public bool Equals(DxfStyleDefinition? other)
            {
                if (other == null) return false;
                return string.Equals(FillColorHex, other.FillColorHex, StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(FontColorHex, other.FontColorHex, StringComparison.OrdinalIgnoreCase) &&
                       Bold == other.Bold;
            }

            public override bool Equals(object? obj) => Equals(obj as DxfStyleDefinition);

            public override int GetHashCode()
            {
                int hash = 17;
                hash = hash * 31 + (FillColorHex != null ? StringComparer.OrdinalIgnoreCase.GetHashCode(FillColorHex) : 0);
                hash = hash * 31 + (FontColorHex != null ? StringComparer.OrdinalIgnoreCase.GetHashCode(FontColorHex) : 0);
                hash = hash * 31 + Bold.GetHashCode();
                return hash;
            }
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
                Name = GetUniqueSheetName(sheetName),
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
                Name = GetUniqueSheetName(sheetName),
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
                Name = GetUniqueSheetName(sheetName),
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

        #region Fluent Image & Formatting APIs

        /// <summary>
        /// Embeds an image into the specified worksheet using OpenXML DrawingML.
        /// </summary>
        /// <param name="sheetName">Target worksheet name.</param>
        /// <param name="imageData">Raw binary bytes of the image.</param>
        /// <param name="format">Format extension, e.g., "png", "jpg", "jpeg".</param>
        /// <param name="column">1-based column index where top-left corner is anchored (e.g., 1 for Col A).</param>
        /// <param name="row">1-based row index where top-left corner is anchored (e.g., 1 for Row 1).</param>
        /// <param name="widthPx">Display width in pixels.</param>
        /// <param name="heightPx">Display height in pixels.</param>
        /// <param name="name">Optional descriptive name for the image element.</param>
        public ExcelWorkbookBuilder AddImage(string sheetName, byte[] imageData, string format, int column, int row, int widthPx, int heightPx, string? name = null)
        {
            var sheet = GetOrAddSheet(sheetName);
            sheet.Images.Add(new ExcelImage(imageData, format, column, row, widthPx, heightPx, name));
            return this;
        }

        /// <summary>
        /// Embeds a configured <see cref="ExcelImage"/> into the specified worksheet.
        /// </summary>
        public ExcelWorkbookBuilder AddImage(string sheetName, ExcelImage image)
        {
            if (image == null) throw new ArgumentNullException(nameof(image));
            var sheet = GetOrAddSheet(sheetName);
            sheet.Images.Add(image);
            return this;
        }

        /// <summary>
        /// Adds a conditional highlight rule (e.g. GreaterThan, LessThan, Equal, Between) to a cell range.
        /// </summary>
        public ExcelWorkbookBuilder AddHighlightRule(string sheetName, string range, CellRuleOperator op, string value1, string? value2 = null, string? fillColorHex = null, string? fontColorHex = null, bool bold = false)
        {
            var sheet = GetOrAddSheet(sheetName);
            sheet.ConditionalFormatting.Add(new ExcelHighlightRule(range, op, value1, value2, fillColorHex, fontColorHex, bold));
            return this;
        }

        /// <summary>
        /// Adds a 2-color or 3-color gradient heatmap conditional formatting rule to a cell range.
        /// </summary>
        public ExcelWorkbookBuilder AddColorScale(string sheetName, string range, string minColorHex = "FFF8696B", string maxColorHex = "FF63BE7B", string? midColorHex = null)
        {
            var sheet = GetOrAddSheet(sheetName);
            sheet.ConditionalFormatting.Add(new ExcelColorScaleRule(range, minColorHex, maxColorHex, midColorHex));
            return this;
        }

        /// <summary>
        /// Adds an in-cell data bar conditional formatting rule to a cell range.
        /// </summary>
        public ExcelWorkbookBuilder AddDataBar(string sheetName, string range, string colorHex = "FF638EC6")
        {
            var sheet = GetOrAddSheet(sheetName);
            sheet.ConditionalFormatting.Add(new ExcelDataBarRule(range, colorHex));
            return this;
        }

        /// <summary>
        /// Enables Excel AutoFilter on the specified worksheet for a specific range or automatically over the populated table headers.
        /// </summary>
        public ExcelWorkbookBuilder SetAutoFilter(string sheetName, string? range = null)
        {
            var sheet = GetOrAddSheet(sheetName);
            if (!string.IsNullOrEmpty(range))
            {
                sheet.AutoFilterRange = range;
            }
            else
            {
                sheet.AutoFilterEnabled = true;
            }
            return this;
        }

        private WorksheetDefinition GetOrAddSheet(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                sheetName = "Sheet1";

            string sanitized = SanitizeSheetName(sheetName);
            foreach (var s in _sheets)
            {
                if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(s.Name, sanitized, StringComparison.OrdinalIgnoreCase))
                {
                    return s;
                }
            }

            var newSheet = new WorksheetDefinition
            {
                Name = GetUniqueSheetName(sheetName)
            };
            _sheets.Add(newSheet);
            return newSheet;
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

            // Collect DXF styles across all sheets
            var dxfs = CollectDxfStyles(_sheets);

            // 1. [Content_Types].xml
            CreateContentTypesEntry(zip, _sheets);

            // 2. _rels/.rels
            CreateGlobalRelsEntry(zip);

            // 3. xl/workbook.xml
            CreateWorkbookEntry(zip, _sheets);

            // 4. xl/_rels/workbook.xml.rels
            CreateWorkbookRelsEntry(zip, _sheets.Count);

            // 5. xl/styles.xml
            CreateStylesEntry(zip, dxfs);

            // 6. xl/worksheets/sheet{N}.xml & Drawings
            int globalImageIndex = 1;
            for (int i = 0; i < _sheets.Count; i++)
            {
                int sheetIndex = i + 1;
                var sheet = _sheets[i];

                CreateWorksheetEntry(zip, sheetIndex, sheet, dxfs);

                if (sheet.Images.Count > 0)
                {
                    CreateSheetRelsEntry(zip, sheetIndex);
                    CreateDrawingEntry(zip, sheetIndex, sheet.Images);
                    CreateDrawingRelsAndMediaEntries(zip, sheetIndex, sheet.Images, ref globalImageIndex);
                }
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

        private static void CreateContentTypesEntry(ZipArchive zip, List<WorksheetDefinition> sheets)
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

            writer.WriteStartElement("Default");
            writer.WriteAttributeString("Extension", "png");
            writer.WriteAttributeString("ContentType", "image/png");
            writer.WriteEndElement();

            writer.WriteStartElement("Default");
            writer.WriteAttributeString("Extension", "jpg");
            writer.WriteAttributeString("ContentType", "image/jpeg");
            writer.WriteEndElement();

            writer.WriteStartElement("Default");
            writer.WriteAttributeString("Extension", "jpeg");
            writer.WriteAttributeString("ContentType", "image/jpeg");
            writer.WriteEndElement();

            writer.WriteStartElement("Override");
            writer.WriteAttributeString("PartName", "/xl/workbook.xml");
            writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            writer.WriteEndElement();

            writer.WriteStartElement("Override");
            writer.WriteAttributeString("PartName", "/xl/styles.xml");
            writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            writer.WriteEndElement();

            for (int i = 1; i <= sheets.Count; i++)
            {
                writer.WriteStartElement("Override");
                writer.WriteAttributeString("PartName", $"/xl/worksheets/sheet{i}.xml");
                writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                writer.WriteEndElement();

                if (sheets[i - 1].Images.Count > 0)
                {
                    writer.WriteStartElement("Override");
                    writer.WriteAttributeString("PartName", $"/xl/drawings/drawing{i}.xml");
                    writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml");
                    writer.WriteEndElement();
                }
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

        private static void CreateStylesEntry(ZipArchive zip, List<DxfStyleDefinition> dxfs)
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

            // dxfs (Differential formatting styles for Conditional Formatting)
            WriteDxfs(writer, dxfs);

            writer.WriteEndElement(); // styleSheet
            writer.WriteEndDocument();
        }

        private static void WriteDxfs(XmlWriter writer, List<DxfStyleDefinition> dxfs)
        {
            if (dxfs == null || dxfs.Count == 0) return;

            writer.WriteStartElement("dxfs");
            writer.WriteAttributeString("count", dxfs.Count.ToString(CultureInfo.InvariantCulture));

            foreach (var dxf in dxfs)
            {
                writer.WriteStartElement("dxf");

                // 1. Font
                if (!string.IsNullOrEmpty(dxf.FontColorHex) || dxf.Bold)
                {
                    writer.WriteStartElement("font");
                    if (dxf.Bold)
                    {
                        writer.WriteStartElement("b");
                        writer.WriteEndElement();
                    }
                    if (!string.IsNullOrEmpty(dxf.FontColorHex))
                    {
                        writer.WriteStartElement("color");
                        writer.WriteAttributeString("rgb", NormalizeColor(dxf.FontColorHex!));
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement(); // font
                }

                // 2. Fill
                if (!string.IsNullOrEmpty(dxf.FillColorHex))
                {
                    writer.WriteStartElement("fill");
                    writer.WriteStartElement("patternFill");
                    writer.WriteStartElement("bgColor");
                    writer.WriteAttributeString("rgb", NormalizeColor(dxf.FillColorHex!));
                    writer.WriteEndElement(); // bgColor
                    writer.WriteEndElement(); // patternFill
                    writer.WriteEndElement(); // fill
                }

                writer.WriteEndElement(); // dxf
            }

            writer.WriteEndElement(); // dxfs
        }

        private void CreateWorksheetEntry(ZipArchive zip, int sheetIndex, WorksheetDefinition sheet, List<DxfStyleDefinition> dxfs)
        {
            var entry = zip.CreateEntry($"xl/worksheets/sheet{sheetIndex}.xml", CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = CreateXmlWriter(stream);

            writer.WriteStartDocument(true);
            writer.WriteStartElement("worksheet", NsSpreadsheet);
            writer.WriteAttributeString("xmlns", "r", null, NsRelationships);

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

            // 3. AutoFilter
            if (!string.IsNullOrEmpty(sheet.AutoFilterRange))
            {
                writer.WriteStartElement("autoFilter");
                writer.WriteAttributeString("ref", sheet.AutoFilterRange);
                writer.WriteEndElement();
            }
            else if (sheet.AutoFilterEnabled && sheet.Headers != null && sheet.Headers.Count > 0)
            {
                string lastCol = ExcelCellAddress.IndexToColumnName(sheet.Headers.Count);
                int lastRow = Math.Max(1, sheet.Rows.Count + 1);
                writer.WriteStartElement("autoFilter");
                writer.WriteAttributeString("ref", $"A1:{lastCol}{lastRow}");
                writer.WriteEndElement();
            }

            // 4. Conditional Formatting
            if (sheet.ConditionalFormatting.Count > 0)
            {
                int priority = 1;
                foreach (var rule in sheet.ConditionalFormatting)
                {
                    writer.WriteStartElement("conditionalFormatting");
                    writer.WriteAttributeString("sqref", rule.Range);

                    if (rule is ExcelHighlightRule hr)
                    {
                        writer.WriteStartElement("cfRule");
                        writer.WriteAttributeString("type", "cellIs");
                        writer.WriteAttributeString("operator", GetOperatorAttribute(hr.Operator));
                        writer.WriteAttributeString("priority", priority.ToString(CultureInfo.InvariantCulture));

                        int dxfId = GetDxfId(dxfs, hr);
                        if (dxfId >= 0)
                        {
                            writer.WriteAttributeString("dxfId", dxfId.ToString(CultureInfo.InvariantCulture));
                        }

                        writer.WriteElementString("formula", NsSpreadsheet, hr.Formula1);
                        if (!string.IsNullOrEmpty(hr.Formula2))
                        {
                            writer.WriteElementString("formula", NsSpreadsheet, hr.Formula2);
                        }

                        writer.WriteEndElement(); // cfRule
                    }
                    else if (rule is ExcelColorScaleRule csr)
                    {
                        writer.WriteStartElement("cfRule");
                        writer.WriteAttributeString("type", "colorScale");
                        writer.WriteAttributeString("priority", priority.ToString(CultureInfo.InvariantCulture));

                        writer.WriteStartElement("colorScale");

                        writer.WriteStartElement("cfvo");
                        writer.WriteAttributeString("type", "min");
                        writer.WriteEndElement();

                        if (!string.IsNullOrEmpty(csr.MidColorHex))
                        {
                            writer.WriteStartElement("cfvo");
                            writer.WriteAttributeString("type", "percentile");
                            writer.WriteAttributeString("val", "50");
                            writer.WriteEndElement();
                        }

                        writer.WriteStartElement("cfvo");
                        writer.WriteAttributeString("type", "max");
                        writer.WriteEndElement();

                        writer.WriteStartElement("color");
                        writer.WriteAttributeString("rgb", NormalizeColor(csr.MinColorHex));
                        writer.WriteEndElement();

                        if (!string.IsNullOrEmpty(csr.MidColorHex))
                        {
                            writer.WriteStartElement("color");
                            writer.WriteAttributeString("rgb", NormalizeColor(csr.MidColorHex!));
                            writer.WriteEndElement();
                        }

                        writer.WriteStartElement("color");
                        writer.WriteAttributeString("rgb", NormalizeColor(csr.MaxColorHex));
                        writer.WriteEndElement();

                        writer.WriteEndElement(); // colorScale
                        writer.WriteEndElement(); // cfRule
                    }
                    else if (rule is ExcelDataBarRule dbr)
                    {
                        writer.WriteStartElement("cfRule");
                        writer.WriteAttributeString("type", "dataBar");
                        writer.WriteAttributeString("priority", priority.ToString(CultureInfo.InvariantCulture));

                        writer.WriteStartElement("dataBar");

                        writer.WriteStartElement("cfvo");
                        writer.WriteAttributeString("type", "min");
                        writer.WriteEndElement();

                        writer.WriteStartElement("cfvo");
                        writer.WriteAttributeString("type", "max");
                        writer.WriteEndElement();

                        writer.WriteStartElement("color");
                        writer.WriteAttributeString("rgb", NormalizeColor(dbr.ColorHex));
                        writer.WriteEndElement();

                        writer.WriteEndElement(); // dataBar
                        writer.WriteEndElement(); // cfRule
                    }

                    writer.WriteEndElement(); // conditionalFormatting
                    priority++;
                }
            }

            // 5. Drawing Reference
            if (sheet.Images.Count > 0)
            {
                writer.WriteStartElement("drawing");
                writer.WriteAttributeString("id", NsRelationships, "rIdDrawing");
                writer.WriteEndElement();
            }

            writer.WriteEndElement(); // worksheet
            writer.WriteEndDocument();
        }

        private static void CreateSheetRelsEntry(ZipArchive zip, int sheetIndex)
        {
            var entry = zip.CreateEntry($"xl/worksheets/_rels/sheet{sheetIndex}.xml.rels", CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = CreateXmlWriter(stream);

            writer.WriteStartDocument(true);
            writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");

            writer.WriteStartElement("Relationship");
            writer.WriteAttributeString("Id", "rIdDrawing");
            writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing");
            writer.WriteAttributeString("Target", $"../drawings/drawing{sheetIndex}.xml");
            writer.WriteEndElement();

            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        private static void CreateDrawingEntry(ZipArchive zip, int sheetIndex, List<ExcelImage> images)
        {
            var entry = zip.CreateEntry($"xl/drawings/drawing{sheetIndex}.xml", CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = CreateXmlWriter(stream);

            writer.WriteStartDocument(true);
            writer.WriteStartElement("xdr", "wsDr", NsDrawingSheet);
            writer.WriteAttributeString("xmlns", "a", null, NsDrawingMain);
            writer.WriteAttributeString("xmlns", "r", null, NsRelationships);

            for (int j = 0; j < images.Count; j++)
            {
                var img = images[j];
                int col = Math.Max(0, img.ColumnIndex - 1);
                int row = Math.Max(0, img.RowIndex - 1);
                long colOff = (long)img.ColOffsetPx * 9525L;
                long rowOff = (long)img.RowOffsetPx * 9525L;
                long cx = (long)img.WidthPx * 9525L;
                long cy = (long)img.HeightPx * 9525L;

                writer.WriteStartElement("xdr", "oneCellAnchor", NsDrawingSheet);

                // <xdr:from>
                writer.WriteStartElement("xdr", "from", NsDrawingSheet);
                writer.WriteElementString("xdr", "col", NsDrawingSheet, col.ToString(CultureInfo.InvariantCulture));
                writer.WriteElementString("xdr", "colOff", NsDrawingSheet, colOff.ToString(CultureInfo.InvariantCulture));
                writer.WriteElementString("xdr", "row", NsDrawingSheet, row.ToString(CultureInfo.InvariantCulture));
                writer.WriteElementString("xdr", "rowOff", NsDrawingSheet, rowOff.ToString(CultureInfo.InvariantCulture));
                writer.WriteEndElement(); // from

                // <xdr:ext cx="..." cy="..."/>
                writer.WriteStartElement("xdr", "ext", NsDrawingSheet);
                writer.WriteAttributeString("cx", cx.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("cy", cy.ToString(CultureInfo.InvariantCulture));
                writer.WriteEndElement(); // ext

                // <xdr:pic>
                writer.WriteStartElement("xdr", "pic", NsDrawingSheet);

                // nvPicPr
                writer.WriteStartElement("xdr", "nvPicPr", NsDrawingSheet);
                writer.WriteStartElement("xdr", "cNvPr", NsDrawingSheet);
                writer.WriteAttributeString("id", (j + 2).ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("name", string.IsNullOrWhiteSpace(img.Name) ? $"Picture {j + 1}" : SanitizeXmlText(img.Name));
                writer.WriteEndElement(); // cNvPr

                writer.WriteStartElement("xdr", "cNvPicPr", NsDrawingSheet);
                writer.WriteStartElement("a", "picLocks", NsDrawingMain);
                writer.WriteAttributeString("noChangeAspect", "1");
                writer.WriteEndElement(); // picLocks
                writer.WriteEndElement(); // cNvPicPr
                writer.WriteEndElement(); // nvPicPr

                // blipFill
                writer.WriteStartElement("xdr", "blipFill", NsDrawingSheet);
                writer.WriteStartElement("a", "blip", NsDrawingMain);
                writer.WriteAttributeString("embed", NsRelationships, $"rId{j + 1}");
                writer.WriteEndElement(); // blip

                writer.WriteStartElement("a", "stretch", NsDrawingMain);
                writer.WriteStartElement("a", "fillRect", NsDrawingMain);
                writer.WriteEndElement(); // fillRect
                writer.WriteEndElement(); // stretch
                writer.WriteEndElement(); // blipFill

                // spPr
                writer.WriteStartElement("xdr", "spPr", NsDrawingSheet);
                writer.WriteStartElement("a", "xfrm", NsDrawingMain);
                writer.WriteStartElement("a", "off", NsDrawingMain);
                writer.WriteAttributeString("x", "0");
                writer.WriteAttributeString("y", "0");
                writer.WriteEndElement(); // off

                writer.WriteStartElement("a", "ext", NsDrawingMain);
                writer.WriteAttributeString("cx", cx.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("cy", cy.ToString(CultureInfo.InvariantCulture));
                writer.WriteEndElement(); // ext
                writer.WriteEndElement(); // xfrm

                writer.WriteStartElement("a", "prstGeom", NsDrawingMain);
                writer.WriteAttributeString("prst", "rect");
                writer.WriteStartElement("a", "avLst", NsDrawingMain);
                writer.WriteEndElement(); // avLst
                writer.WriteEndElement(); // prstGeom
                writer.WriteEndElement(); // spPr

                writer.WriteEndElement(); // pic

                writer.WriteStartElement("xdr", "clientData", NsDrawingSheet);
                writer.WriteEndElement(); // clientData

                writer.WriteEndElement(); // oneCellAnchor
            }

            writer.WriteEndElement(); // wsDr
            writer.WriteEndDocument();
        }

        private static void CreateDrawingRelsAndMediaEntries(ZipArchive zip, int sheetIndex, List<ExcelImage> images, ref int globalImageIndex)
        {
            int startGlobalIndex = globalImageIndex;

            // 1. Write the drawing relationship XML part
            var relsEntry = zip.CreateEntry($"xl/drawings/_rels/drawing{sheetIndex}.xml.rels", CompressionLevel.Fastest);
            using (var relsStream = relsEntry.Open())
            using (var writer = CreateXmlWriter(relsStream))
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");

                for (int j = 0; j < images.Count; j++)
                {
                    var img = images[j];
                    int currentGlobalIndex = startGlobalIndex + j;
                    string ext = string.IsNullOrWhiteSpace(img.Format) ? "png" : img.Format.ToLowerInvariant();

                    writer.WriteStartElement("Relationship");
                    writer.WriteAttributeString("Id", $"rId{j + 1}");
                    writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                    writer.WriteAttributeString("Target", $"../media/image{currentGlobalIndex}.{ext}");
                    writer.WriteEndElement();
                }

                writer.WriteEndElement(); // Relationships
                writer.WriteEndDocument();
            }

            // 2. Write image binaries sequentially after closing the relationship stream
            for (int j = 0; j < images.Count; j++)
            {
                var img = images[j];
                int currentGlobalIndex = startGlobalIndex + j;
                string ext = string.IsNullOrWhiteSpace(img.Format) ? "png" : img.Format.ToLowerInvariant();
                string mediaPath = $"xl/media/image{currentGlobalIndex}.{ext}";

                var mediaEntry = zip.CreateEntry(mediaPath, CompressionLevel.Optimal);
                using var mediaStream = mediaEntry.Open();
                mediaStream.Write(img.Data, 0, img.Data.Length);
            }

            globalImageIndex = startGlobalIndex + images.Count;
        }

        private static List<DxfStyleDefinition> CollectDxfStyles(List<WorksheetDefinition> sheets)
        {
            var dxfs = new List<DxfStyleDefinition>();
            foreach (var sheet in sheets)
            {
                foreach (var rule in sheet.ConditionalFormatting)
                {
                    if (rule is ExcelHighlightRule hr && (!string.IsNullOrEmpty(hr.FillColorHex) || !string.IsNullOrEmpty(hr.FontColorHex) || hr.Bold))
                    {
                        var def = new DxfStyleDefinition
                        {
                            FillColorHex = hr.FillColorHex,
                            FontColorHex = hr.FontColorHex,
                            Bold = hr.Bold
                        };
                        if (!dxfs.Contains(def))
                        {
                            dxfs.Add(def);
                        }
                    }
                }
            }
            return dxfs;
        }

        private static int GetDxfId(List<DxfStyleDefinition> dxfs, ExcelHighlightRule hr)
        {
            if (string.IsNullOrEmpty(hr.FillColorHex) && string.IsNullOrEmpty(hr.FontColorHex) && !hr.Bold)
                return -1;

            for (int i = 0; i < dxfs.Count; i++)
            {
                var d = dxfs[i];
                if (string.Equals(d.FillColorHex, hr.FillColorHex, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(d.FontColorHex, hr.FontColorHex, StringComparison.OrdinalIgnoreCase) &&
                    d.Bold == hr.Bold)
                {
                    return i;
                }
            }
            return -1;
        }

        private static string GetOperatorAttribute(CellRuleOperator op) => op switch
        {
            CellRuleOperator.GreaterThan => "greaterThan",
            CellRuleOperator.LessThan => "lessThan",
            CellRuleOperator.Equal => "equal",
            CellRuleOperator.NotEqual => "notEqual",
            CellRuleOperator.GreaterThanOrEqual => "greaterThanOrEqual",
            CellRuleOperator.LessThanOrEqual => "lessThanOrEqual",
            CellRuleOperator.Between => "between",
            CellRuleOperator.NotBetween => "notBetween",
            _ => "greaterThan"
        };

        private static string NormalizeColor(string hex)
        {
            if (string.IsNullOrWhiteSpace(hex)) return "FF000000";
            string clean = hex.Trim().TrimStart('#');
            if (clean.Length == 6)
            {
                return "FF" + clean.ToUpperInvariant();
            }
            if (clean.Length == 8)
            {
                return clean.ToUpperInvariant();
            }
            return "FF000000";
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

        private string GetUniqueSheetName(string rawName)
        {
            string baseName = SanitizeSheetName(rawName);
            string uniqueName = baseName;
            int counter = 1;

            while (ContainsSheetName(uniqueName))
            {
                string suffix = $"_{counter++}";
                if (baseName.Length + suffix.Length > 31)
                {
                    uniqueName = baseName.Substring(0, 31 - suffix.Length) + suffix;
                }
                else
                {
                    uniqueName = baseName + suffix;
                }
            }

            return uniqueName;
        }

        private bool ContainsSheetName(string name)
        {
            for (int i = 0; i < _sheets.Count; i++)
            {
                if (string.Equals(_sheets[i].Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
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
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];

                // Preserve valid surrogate pairs (e.g., emojis, supplementary multilingual characters)
                if (char.IsHighSurrogate(ch) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    sb.Append(ch);
                    sb.Append(text[++i]);
                    continue;
                }

                // XML 1.0 valid single-char code points:
                // #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD]
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
