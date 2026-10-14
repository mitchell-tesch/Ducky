using System;
using DuckDB.NET.Data;

namespace GhDucky.Utils
{
    /// <summary>
    /// Builds per-column delegates that append a single cell through DuckDB's
    /// high-performance Appender API.
    /// <para>
    /// Resolving the SQL type once per column and returning a closed-over
    /// delegate avoids re-running a type switch for every individual cell,
    /// which matters when appending tens of thousands of rows.
    /// </para>
    /// </summary>
    internal static class DuckDbAppenderFactory
    {
        /// <summary>
        /// Returns a delegate that appends one value in the form the appender
        /// expects for <paramref name="sqlType"/>, writing SQL NULL for nulls.
        /// </summary>
        public static Action<IDuckDBAppenderRow, object> BuildCellAppender(DuckyColumnType sqlType)
        {
            switch (sqlType)
            {
                case DuckyColumnType.Boolean:
                    return (row, value) =>
                    {
                        if (value is null) { row.AppendNullValue(); return; }
                        row.AppendValue((bool?)TypeMapping.CoerceForAppender(value, DuckyColumnType.Boolean));
                    };
                case DuckyColumnType.Integer:
                    return (row, value) =>
                    {
                        if (value is null) { row.AppendNullValue(); return; }
                        row.AppendValue((int?)TypeMapping.CoerceForAppender(value, DuckyColumnType.Integer));
                    };
                case DuckyColumnType.BigInt:
                    return (row, value) =>
                    {
                        if (value is null) { row.AppendNullValue(); return; }
                        row.AppendValue((long?)TypeMapping.CoerceForAppender(value, DuckyColumnType.BigInt));
                    };
                case DuckyColumnType.Double:
                    return (row, value) =>
                    {
                        if (value is null) { row.AppendNullValue(); return; }
                        row.AppendValue((double?)TypeMapping.CoerceForAppender(value, DuckyColumnType.Double));
                    };
                case DuckyColumnType.Timestamp:
                    return (row, value) =>
                    {
                        if (value is null) { row.AppendNullValue(); return; }
                        row.AppendValue((DateTime?)TypeMapping.CoerceForAppender(value, DuckyColumnType.Timestamp));
                    };
                case DuckyColumnType.Varchar:
                default:
                    return (row, value) =>
                    {
                        if (value is null) { row.AppendNullValue(); return; }
                        row.AppendValue((string)TypeMapping.CoerceForAppender(value, DuckyColumnType.Varchar));
                    };
            }
        }

        /// <summary>
        /// Ensures column names are unique and non-empty, which
        /// <c>CREATE TABLE</c> requires. Duplicates gain a numeric suffix.
        /// </summary>
        public static System.Collections.Generic.List<string> MakeUniqueColumnNames(
            System.Collections.Generic.IReadOnlyList<string> names)
        {
            var seen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new System.Collections.Generic.List<string>(names.Count);

            for (var i = 0; i < names.Count; i++)
            {
                var candidate = string.IsNullOrWhiteSpace(names[i])
                    ? "col_" + i.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : names[i].Trim();

                var unique = candidate;
                var suffix = 2;
                while (!seen.Add(unique))
                {
                    unique = candidate + "_" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    suffix++;
                }
                result.Add(unique);
            }

            return result;
        }
    }
}

