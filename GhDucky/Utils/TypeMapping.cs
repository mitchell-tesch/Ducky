using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Globalization;
using System.Numerics;
using System.Text;
using Grasshopper.Kernel.Types;

namespace GhDucky.Utils
{
    /// Maps between DuckDB-native values, .NET CLR values and Grasshopper goo.
    /// Read path: unbox reader output -> IGH_Goo.
    /// Write path: unwrap IGH_Goo / GH primitives -> CLR values appendable through DuckDBAppender.
    internal static class TypeMapping
    {
        /// Wraps a value coming out of a DuckDBDataReader cell into a Grasshopper-native goo.
        /// When <paramref name="typed"/> is true (the default), temporal values surface as
        /// <see cref="GH_Time"/> so downstream date arithmetic works; when false, all
        /// non-primitive types are stringified (legacy behaviour, useful for diffing).
        public static IGH_Goo ToGoo(object value, bool typed = true)
        {
            // Return null for SQL NULL so downstream components can distinguish
            // "missing value" from "empty string". Grasshopper interprets a null
            // goo as a gap in the data tree, which matches DuckDB NULL semantics.
            if (value is null || value is DBNull)
                return null;

            switch (value)
            {
                case bool b:
                    return new GH_Boolean(b);
                case sbyte sb:
                    return new GH_Integer(sb);
                case byte by:
                    return new GH_Integer(by);
                case short s:
                    return new GH_Integer(s);
                case ushort us:
                    return new GH_Integer(us);
                case int i:
                    return new GH_Integer(i);
                case uint ui:
                    return ui <= int.MaxValue
                        ? new GH_Integer((int)ui)
                        : new GH_Number(ui);
                case long l:
                    return l is >= int.MinValue and <= int.MaxValue
                        ? new GH_Integer((int)l)
                        : new GH_Number(l);
                case ulong ul:
                    return ul <= int.MaxValue
                        ? new GH_Integer((int)ul)
                        : new GH_Number(ul);
                case float f:
                    return new GH_Number(f);
                case double d:
                    return new GH_Number(d);
                case decimal dec:
                    return new GH_Number((double)dec);
                case string str:
                    return new GH_String(str);
                case Guid g:
                    return new GH_String(g.ToString("D"));
                case DateTime dt:
                    return typed
                        ? new GH_Time(dt)
                        : new GH_String(dt.ToString("o", CultureInfo.InvariantCulture));
                case DateTimeOffset dto:
                    return typed
                        ? new GH_Time(dto.UtcDateTime)
                        : new GH_String(dto.ToString("o", CultureInfo.InvariantCulture));
                case DateOnly d:
                    return typed
                        ? new GH_Time(d.ToDateTime(TimeOnly.MinValue))
                        : new GH_String(d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                case TimeOnly t:
                    return typed
                        ? new GH_Time(new DateTime(1, 1, 1, t.Hour, t.Minute, t.Second, t.Millisecond, DateTimeKind.Unspecified))
                        : new GH_String(t.ToString("HH:mm:ss.fffffff", CultureInfo.InvariantCulture));
                case TimeSpan ts:
                    return new GH_String(ts.ToString("c", CultureInfo.InvariantCulture));
                case byte[] bytes:
                    return new GH_String(Convert.ToBase64String(bytes));
                case BigInteger bi:
                    return new GH_String(bi.ToString(CultureInfo.InvariantCulture));
                case IEnumerable en when !(value is string):
                    return new GH_String(FormatEnumerable(en));
            }

            return new GH_String(Convert.ToString(value, CultureInfo.InvariantCulture));
        }

        private static string FormatEnumerable(IEnumerable en)
        {
            var sb = new StringBuilder();
            sb.Append('[');
            bool first = true;
            foreach (var item in en)
            {
                if (!first) sb.Append(", ");
                first = false;
                sb.Append(Convert.ToString(item, CultureInfo.InvariantCulture));
            }
            sb.Append(']');
            return sb.ToString();
        }

        /// Unwraps a goo into the underlying CLR value usable by the DuckDB appender.
        /// Returns null for null/empty values; the appender will write SQL NULL.
        public static object Unwrap(IGH_Goo goo)
        {
            if (goo is null) return null;

            switch (goo)
            {
                case GH_Boolean b: return b.Value;
                case GH_Integer i: return i.Value;
                case GH_Number n: return n.Value;
                case GH_String s: return s.Value;
                case GH_Time t: return t.Value;
                case GH_Guid g: return g.Value.ToString("D");
            }

            var scripted = goo.ScriptVariable();
            return scripted ?? goo.ToString();
        }

        /// Type-priority for inferring a column type when scanning sample rows.
        /// Higher value = more permissive (wider) type.
        private static int Priority(DuckyColumnType type) => type switch
        {
            DuckyColumnType.Boolean   => 1,
            DuckyColumnType.Integer   => 2,
            DuckyColumnType.BigInt    => 3,
            DuckyColumnType.Double    => 4,
            DuckyColumnType.Timestamp => 5,
            DuckyColumnType.Varchar   => 6,
            _ => 7,
        };

        /// Infers the widest DuckDB column type required to hold every sample value.
        public static DuckyColumnType InferColumnType(IEnumerable<object> samples)
        {
            DuckyColumnType? best = null;
            foreach (var raw in samples)
            {
                if (raw is null) continue;
                var classified = ClassifyOne(raw);
                if (classified is null) continue;

                if (best is null || Priority(classified.Value) > Priority(best.Value))
                    best = classified;
            }
            return best ?? DuckyColumnType.Varchar;
        }

        private static DuckyColumnType? ClassifyOne(object value)
        {
            switch (value)
            {
                case bool _:
                    return DuckyColumnType.Boolean;
                case sbyte _:
                case byte _:
                case short _:
                case ushort _:
                case int _:
                    return DuckyColumnType.Integer;
                case uint _:
                case long _:
                case ulong _:
                    return DuckyColumnType.BigInt;
                case float _:
                case double _:
                case decimal _:
                    return DuckyColumnType.Double;
                case DateTime _:
                case DateTimeOffset _:
                    return DuckyColumnType.Timestamp;
                case string _:
                case Guid _:
                case TimeSpan _:
                    break;
            }

            return DuckyColumnType.Varchar;
        }

        /// Coerces an unwrapped CLR value to the form the appender expects for the given SQL type.
        public static object CoerceForAppender(object value, DuckyColumnType sqlType)
        {
            if (value is null) return null;

            try
            {
                switch (sqlType)
                {
                    case DuckyColumnType.Boolean:
                        if (value is bool b) return b;
                        return Convert.ToBoolean(value, CultureInfo.InvariantCulture);

                    case DuckyColumnType.Integer:
                        return Convert.ToInt32(value, CultureInfo.InvariantCulture);

                    case DuckyColumnType.BigInt:
                        return Convert.ToInt64(value, CultureInfo.InvariantCulture);

                    case DuckyColumnType.Double:
                        return Convert.ToDouble(value, CultureInfo.InvariantCulture);

                    case DuckyColumnType.Timestamp:
                        if (value is DateTime dt) return dt;
                        if (value is DateTimeOffset dto) return dto.UtcDateTime;
                        if (value is string str &&
                            DateTime.TryParse(str, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsed))
                            return parsed;
                        return Convert.ToDateTime(value, CultureInfo.InvariantCulture);

                    case DuckyColumnType.Varchar:
                    default:
                        return Convert.ToString(value, CultureInfo.InvariantCulture);
                }
            }
            catch
            {
                return Convert.ToString(value, CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        /// Resolves column names from a user-supplied list, padding with
        /// "col_0", "col_1", ... for any missing entries.
        /// </summary>
        public static List<string> ResolveColumnNames(IReadOnlyList<string> supplied, int count)
        {
            var names = new List<string>(count);
            for (var i = 0; i < count; i++)
            {
                if (supplied != null && i < supplied.Count && !string.IsNullOrWhiteSpace(supplied[i]))
                    names.Add(supplied[i].Trim());
                else
                    names.Add("col_" + i.ToString(CultureInfo.InvariantCulture));
            }
            return names;
        }

        /// <summary>
        /// Resolves column types from a user-supplied list, auto-inferring from
        /// sample values for any missing entries. Samples up to
        /// <paramref name="maxSampleRows"/> values per column for inference
        /// instead of scanning every value (large columns benefit significantly).
        /// </summary>
        public static List<DuckyColumnType> ResolveColumnTypes(
            IReadOnlyList<string> supplied, object[][] columns, int maxSampleRows = 1024)
        {
            var types = new List<DuckyColumnType>(columns.Length);
            for (var i = 0; i < columns.Length; i++)
            {
                if (supplied != null && i < supplied.Count && !string.IsNullOrWhiteSpace(supplied[i]))
                    types.Add(DuckyColumnTypeExtensions.ParseOrVarchar(supplied[i]));
                else
                    types.Add(InferColumnType(SampleValues(columns[i], maxSampleRows)));
            }
            return types;
        }

        /// <summary>
        /// Returns up to <paramref name="max"/> evenly-spaced samples from the array.
        /// For small arrays the full array is returned (zero allocation).
        /// </summary>
        private static IEnumerable<object> SampleValues(object[] values, int max)
        {
            if (values.Length <= max)
                return values;

            var sampled = new object[max];
            var step = (double)values.Length / max;
            for (var i = 0; i < max; i++)
                sampled[i] = values[(int)(i * step)];
            return sampled;
        }

        /// <summary>
        /// Builds a per-column reader that pulls a typed value out of a
        /// <see cref="DbDataReader"/> and wraps it as a Grasshopper goo,
        /// avoiding the per-cell <c>GetValue</c> boxing and the broad
        /// <see cref="ToGoo"/> switch.
        /// <para>
        /// Falls back to <see cref="ToGoo"/> for SQL types we do not specially
        /// recognise (e.g. arrays, structs, BLOBs, ENUMs).
        /// </para>
        /// </summary>
        /// <param name="dataTypeName">
        /// Value of <see cref="DbDataReader.GetDataTypeName(int)"/> for the column.
        /// Parametric forms like <c>DECIMAL(10,2)</c> are accepted.
        /// </param>
        /// <param name="typed">
        /// When <c>true</c>, temporal columns are surfaced as <see cref="GH_Time"/>;
        /// when <c>false</c>, they are serialised as ISO-8601 strings.
        /// </param>
        public static Func<DbDataReader, int, IGH_Goo> BuildColumnReader(string dataTypeName, bool typed)
        {
            var upper = (dataTypeName ?? string.Empty).Trim().ToUpperInvariant();

            // Strip parametric suffix: "DECIMAL(10,2)" → "DECIMAL".
            var paren = upper.IndexOf('(');
            if (paren > 0) upper = upper.Substring(0, paren);

            switch (upper)
            {
                case "BOOLEAN":
                case "BOOL":
                    return (r, i) => new GH_Boolean(r.GetBoolean(i));

                case "TINYINT":
                case "SMALLINT":
                case "INTEGER":
                case "INT":
                case "INT4":
                case "UTINYINT":
                case "USMALLINT":
                    return (r, i) => new GH_Integer(r.GetInt32(i));

                case "BIGINT":
                case "INT8":
                case "UINTEGER":
                    return (r, i) =>
                    {
                        var v = r.GetInt64(i);
                        return v is >= int.MinValue and <= int.MaxValue
                            ? new GH_Integer((int)v)
                            : new GH_Number(v);
                    };

                case "DOUBLE":
                case "REAL":
                case "FLOAT":
                case "FLOAT4":
                case "FLOAT8":
                case "DECIMAL":
                case "NUMERIC":
                    return (r, i) => new GH_Number(r.GetDouble(i));

                case "VARCHAR":
                case "TEXT":
                case "STRING":
                case "CHAR":
                    return (r, i) => new GH_String(r.GetString(i));

                case "DATE":
                case "TIME":
                case "TIMESTAMP":
                case "DATETIME":
                case "TIMESTAMP_S":
                case "TIMESTAMP_MS":
                case "TIMESTAMP_NS":
                    return typed
                        ? (r, i) => new GH_Time(r.GetDateTime(i))
                        : (r, i) => new GH_String(r.GetDateTime(i).ToString("o", CultureInfo.InvariantCulture));

                case "UUID":
                    return (r, i) => new GH_String(r.GetGuid(i).ToString("D"));

                default:
                    // HUGEINT, BLOB, ENUM, LIST, STRUCT, MAP, etc. — keep the
                    // broad ToGoo path so the existing semantics are preserved.
                    return (r, i) => ToGoo(r.GetValue(i), typed);
            }
        }
    }
}
