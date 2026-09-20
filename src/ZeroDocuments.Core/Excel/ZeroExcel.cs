namespace ZeroDocuments.Excel
{
    /// <summary>
    /// Entry point for ZeroDocuments Excel fluent builder and high-level operations.
    /// </summary>
    public static class ZeroExcel
    {
        /// <summary>
        /// Creates a new fluent Excel workbook builder supporting multiple worksheets.
        /// </summary>
        public static ExcelWorkbookBuilder Create() => new ExcelWorkbookBuilder();
    }
}
