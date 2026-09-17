# ZeroDocuments

[![NuGet Version](https://img.shields.io/badge/nuget-v1.0.0-blue.svg)](https://www.nuget.org/packages/ZeroDocuments.Core/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Zero External Dependencies](https://img.shields.io/badge/Dependencies-0%20External-brightgreen.svg)]()
[![Tests: 24 Passed](https://img.shields.io/badge/Tests-24%20Passed%20(100%25)-brightgreen.svg)]()
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
| **.NET 4.6.2 Compatibility**| Prone to `System.IO.Compression` binding issues | Complex assembly deployment | **Native Auto-Resolver built-in** |
| **Performance** | High memory overhead (DOM tree) | Heavy UI/DOM abstraction | **Stream-based OpenXML parser** |

---

## 🏛️ Comprehensive Architecture

```
ZeroDocuments.Core
 ├── Excel/
 │    ├── ExcelReader.cs                 # High-speed OpenXML stream parser & DataTable reader
 │    ├── ExcelWriter.cs                 # Pure C# .xlsx package builder & exporter
 │    └── Models/
 │         ├── ExcelCellAddress.cs       # Coordinate calculations (e.g., "D24" -> Col 4, Row 24)
 │         ├── ExcelRow.cs               # Column-indexed lightweight row model
 │         └── ExcelCell.cs              # Typed cell value container
 ├── Csv/
 │    ├── CsvReader.cs                   # RFC 4180 streaming CSV parser with delimiter flexibility
 │    └── CsvWriter.cs                   # Fast DataTable & typed collection CSV generator
 └── Common/
      └── RuntimeAssemblyResolver.cs     # Self-healing assembly binder for .NET Framework runtimes
```

---

## 🌟 Key Capabilities

### 1. Pure C# OpenXML Excel Reader (`ExcelReader`)
- **Header-Bounded Range Parsing**: Read data bounded by a specific header range (e.g. `D24:T24`), ideal for complex enterprise invoice and production templates with top/side metadata.
- **Direct Coordinate Parsing**: Extract specific ranges like `A2:G500` or stream entire worksheets row-by-row.
- **Zero Third-Party Dependencies**: Reads `.xlsx` OpenXML packages directly via standard ZIP streaming and XML token parsing.
- **DataTable & Collection Mapping**: Converts rows directly into `System.Data.DataTable` or strongly-typed POCO collections.

### 2. Pure C# OpenXML Excel Writer (`ExcelWriter`)
- **Direct DataTable Export**: Export any `DataTable` directly to a valid `.xlsx` file with auto-generated shared strings, workbook XML, and worksheet parts.
- **Typed Collection Export**: Stream any `IEnumerable<T>` to Excel with automatic column header inference and type formatting.
- **Custom Sheet Names**: Organize reports with custom worksheet naming.

### 3. RFC 4180 CSV Engine (`CsvReader` & `CsvWriter`)
- **Flexible Delimiters**: Support for comma (`,`), semicolon (`;`), tab (`\t`), and custom industrial delimiters.
- **RFC 4180 Compliant**: Handles embedded quotes, line breaks within cells, and special characters cleanly.
- **Bidirectional Conversion**: Convert between CSV files/streams and `DataTable` or `IEnumerable<T>` effortlessly.

### 4. Self-Healing Runtime Assembly Resolver (`RuntimeAssemblyResolver`)
- Automatically resolves `.NET Framework 4.6.2` assembly binding redirects for `System.IO.Compression`, preventing the classic `FileNotFoundException` or `FileLoadException` on legacy client deployments.

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

### 3. Export DataTable & Collections to Excel (.xlsx)
```csharp
using ZeroDocuments.Excel;

// Export DataTable directly
DataTable inventoryTable = GetInventoryData();
ExcelWriter.WriteToFile("InventoryExport.xlsx", inventoryTable, sheetName: "CurrentStock");

// Export strongly-typed DTO list
var products = new List<ProductDto>
{
    new ProductDto { Id = 1, Sku = "SKU-001", Price = 25.5m },
    new ProductDto { Id = 2, Sku = "SKU-002", Price = 42.0m }
};

ExcelWriter.WriteToFile("Products.xlsx", products, sheetName: "ProductCatalog");
```

### 4. Read and Write CSV Files
```csharp
using ZeroDocuments.Csv;

// Read CSV to DataTable
DataTable csvData = CsvReader.ReadToDataTable("telemetry.csv", delimiter: ',');

// Write DataTable to CSV
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
