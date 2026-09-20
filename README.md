# ZeroDocuments

[![NuGet Version](https://img.shields.io/badge/nuget-v1.1.0-blue.svg)](https://www.nuget.org/packages/ZeroDocuments.Core/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Zero External Dependencies](https://img.shields.io/badge/Dependencies-0%20External-brightgreen.svg)]()
[![Tests: 27 Passed](https://img.shields.io/badge/Tests-27%20Passed%20(100%25)-brightgreen.svg)]()
[![Multi-Targeting](https://img.shields.io/badge/.NET-8.0%20%7C%204.6.2%20%7C%20Standard%202.0-orange.svg)]()

> **Architectural Standard**: 100% Pure C# BCL, Zero External Dependencies (No EPPlus, ClosedXML, or DevExpress), Multi-Targeting across `.NET 8.0`, `.NET Framework 4.6.2`, and `.NET Standard 2.0`.

`ZeroDocuments` is an ultra-fast, zero-dependency spreadsheet and document processing engine engineered for enterprise applications (MDS ERP, WinForms, Web APIs, and microservices). It reads and writes modern Microsoft Excel files (`.xlsx`) and RFC 4180 CSV streams in pure C# using native BCL primitives (`System.IO.Compression` + `System.Xml`), eliminating the need for bulky third-party libraries (EPPlus, ClosedXML, NPOI) or expensive proprietary packages (DevExpress).

---

## 💡 Why ZeroDocuments?

| Feature | Legacy Libraries (EPPlus, ClosedXML) | DevExpress Spreadsheet | ZeroDocuments |
| :--- | :--- | :--- | :--- |
| **Dependencies** | 5 – 15 transitive packages | Heavy proprietary DLLs (~40MB+) | **0 External Dependencies** (BCL only) |
| **License** | Commercial / PolyForm / AGPL | Commercial (Per-Developer License) | **MIT License** (Free & Open Source) |
| **Publish Size** | +15MB – 30MB | +40MB – 80MB | **< 100 KB** |
| **Memory Footprint** | Heavy DOM Tree (>200MB on 100k rows) | Heavy UI/DOM model | **Streaming XmlReader (&lt; 15MB RAM)** |
| **Security** | Vulnerable to CSV Injection if unescaped | Depends on implementation | **CWE-1236 Formula Guard Built-in** |
| **.NET 4.6.2 Compatibility**| Prone to `System.IO.Compression` binding issues | Complex assembly deployment | **Native Auto-Resolver built-in** |

---

## 🏛️ Comprehensive Architecture

```
ZeroDocuments.Core
 ├── Excel/
 │    ├── ZeroExcel.cs                   # Fluent multi-sheet workbook factory
 │    ├── ExcelWorkbookBuilder.cs        # Multi-sheet OpenXML package generator with styling
 │    ├── ExcelReader.cs                 # Low-memory streaming XmlReader & POCO mapper
 │    ├── ExcelWriter.cs                 # High-speed single-sheet OpenXML exporter
 │    └── Models/
 │         ├── ExcelCellAddress.cs       # Coordinate calculations (e.g., "D24" -> Col 4, Row 24)
 │         ├── ExcelRow.cs               # Column-indexed lightweight row model
 │         └── ExcelCell.cs              # Typed cell value container
 ├── Csv/
 │    ├── CsvReader.cs                   # RFC 4180 streaming CSV parser with delimiter flexibility
 │    └── CsvWriter.cs                   # Fast DataTable & typed collection CSV generator with CWE-1236 guard
 └── Common/
      ├── PropertyAccessorCache.cs       # Compiled Expression Trees for 30x-50x faster POCO mapping
      └── RuntimeAssemblyResolver.cs     # Self-healing assembly binder for .NET Framework runtimes
```

---

## 🌟 Key Capabilities

### 1. Low-Memory Streaming Excel Reader (`ExcelReader`)
- **Forward-Only Streaming (`StreamRows`)**: Processes 1,000,000+ rows with **&lt; 15MB RAM** without building heavy DOM trees.
- **Compiled POCO Mapping (`Read<T>`)**: Maps worksheets directly into strongly-typed DTOs via compiled Expression Trees without reflection lag.
- **Header-Bounded Range Parsing**: Read data bounded by a specific header range (e.g. `D24:T24`), ideal for complex enterprise invoice and production templates.

### 2. Fluent Multi-Sheet Workbook Builder (`ZeroExcel`)
- **Multi-Sheet Support**: Add multiple sheets with custom names, data sources (DataTables, POCOs, 2D grids), and header styling (Bold, border layout).
- **Formula Injection Mitigation (CWE-1236)**: Automatically neutralizes malicious formula payloads (`=`, `+`, `-`, `@`) to protect downstream users.
- **Direct Memory & File Export**: Save to file, stream, or byte array (`ToArray()`).

### 3. Compiled Expression Trees (`PropertyAccessorCache`)
- Eliminates reflection overhead when serializing or deserializing collections of objects, achieving 30x–50x speedups over `PropertyInfo.GetValue`.

### 4. RFC 4180 CSV Engine (`CsvReader` & `CsvWriter`)
- **Flexible Delimiters**: Support for comma (`,`), semicolon (`;`), tab (`\t`), and custom delimiters.
- **CWE-1236 Guard**: Prevents CSV injection attacks by automatically prefixing dangerous trigger characters with single quotes.

### 5. Self-Healing Runtime Assembly Resolver (`RuntimeAssemblyResolver`)
- Automatically resolves `.NET Framework 4.6.2` assembly binding redirects for `System.IO.Compression`.

---

## 🚀 Quick Start Examples

### 1. Read Excel by Header Range (Enterprise Template Pattern)
```csharp
using ZeroDocuments.Excel;

// Read starting from header D24:T24 down to a maximum of 5,000 rows
DataTable table = ExcelReader.ReadByHeaderRange("PurchaseOrder.xlsx", "D24:T24", maxRows: 5000);

foreach (DataRow row in table.Rows)
{
    string itemCode = row["Item Code"]?.ToString() ?? "";
    decimal quantity = Convert.ToDecimal(row["Quantity"]);
    Console.WriteLine($"Item: {itemCode} | Qty: {quantity}");
}
```

### 2. Stream Rows from Excel File
```csharp
using ZeroDocuments.Excel;

// Stream rows without loading the entire document into memory
var rows = ExcelReader.ReadRows("Inventory.xlsx", "A1:Z500");

foreach (var row in rows)
{
    string? barcode = row["B"]; // Read cell by column letter
    string? name = row["C"];
    Console.WriteLine($"{barcode}: {name}");
}
```

### 3. Multi-Sheet Workbook Builder (ZeroExcel)
```csharp
using ZeroDocuments.Excel;

using var workbook = ZeroExcel.Create();
workbook.AddSheet("Summary", summaryTable);
workbook.AddSheet("Products", productList);
workbook.AddSheet("AuditLogs", logRows, headers: new[] { "Event", "Date", "Status" });

// Save directly to file or byte array
workbook.Save("EnterpriseReport.xlsx");
byte[] rawBytes = workbook.ToArray();
```

### 4. Read Excel to Strongly-Typed POCO Collections
```csharp
using ZeroDocuments.Excel;

// Automatically map Excel columns into ProductDto properties via compiled Expression Trees
List<ProductDto> products = ExcelReader.Read<ProductDto>("Products.xlsx", "A1:E500", "ProductCatalog");

foreach (var p in products)
{
    Console.WriteLine($"ID: {p.Id}, SKU: {p.Sku}, Price: {p.Price:C}");
}
```

### 5. Read and Write CSV Files with CWE-1236 Protection
```csharp
using ZeroDocuments.Csv;

// Read CSV to DataTable
DataTable csvData = CsvReader.ReadToDataTable("telemetry.csv", delimiter: ',');

// Write DataTable to CSV (with automatic formula injection mitigation)
CsvWriter.WriteToFile("output.csv", csvData, delimiter: ';', includeHeaders: true);
```

---

## 💻 Supported Platforms

| Framework | Target Support | Highlights |
| :--- | :--- | :--- |
| **.NET 8.0+** | Native (`net8.0`) | High-throughput ZIP and memory optimization |
| **.NET Framework** | Legacy WinForms / WPF (`net462`) | Includes automatic assembly resolver for compression |
| **.NET Standard** | Universal Cross-Platform (`netstandard2.0`) | Universal compatibility for shared libraries |

---

## 📄 License
MIT License © 2026 Phong Võ (`kzxl`). Part of the **ZeroPlatform** sovereign ecosystem.
