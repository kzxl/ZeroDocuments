using System;
using System.Collections.Generic;
using System.Data;

namespace ZeroDocuments.Common
{
    /// <summary>
    /// Adapts DataTables and POCO collections into lazily-enumerated row sequences shared by the Excel and CSV writers.
    /// </summary>
    internal static class TabularSource
    {
        public static IReadOnlyList<string> GetHeaders(DataTable table)
        {
            var headers = new string[table.Columns.Count];
            for (int i = 0; i < headers.Length; i++)
            {
                headers[i] = table.Columns[i].ColumnName;
            }
            return headers;
        }

        public static IReadOnlyList<string> GetHeaders(PropertyAccessorCache.PropertyAccessorInfo[] accessors)
        {
            var headers = new string[accessors.Length];
            for (int i = 0; i < accessors.Length; i++)
            {
                headers[i] = accessors[i].Name;
            }
            return headers;
        }

        /// <summary>
        /// Yields DataTable rows as value arrays. DBNull becomes null; rows marked Deleted are skipped
        /// (accessing them would throw <see cref="DeletedRowInaccessibleException"/>).
        /// </summary>
        public static IEnumerable<IReadOnlyList<object?>> FromDataTable(DataTable table)
        {
            int columnCount = table.Columns.Count;
            foreach (DataRow row in table.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;

                var values = new object?[columnCount];
                for (int i = 0; i < columnCount; i++)
                {
                    object v = row[i];
                    values[i] = v == DBNull.Value ? null : v;
                }
                yield return values;
            }
        }

        /// <summary>
        /// Yields POCO property values via compiled getters. Null items are skipped.
        /// </summary>
        public static IEnumerable<IReadOnlyList<object?>> FromObjects<T>(IEnumerable<T> data, PropertyAccessorCache.PropertyAccessorInfo[] accessors)
        {
            foreach (var item in data)
            {
                if (item == null) continue;

                var values = new object?[accessors.Length];
                for (int i = 0; i < accessors.Length; i++)
                {
                    values[i] = accessors[i].Getter(item);
                }
                yield return values;
            }
        }
    }
}
