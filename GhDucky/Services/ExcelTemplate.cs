using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MiniExcelLib.OpenXml;

namespace GhDucky.Services
{
    /// <summary>
    /// Fills a styled Excel template with data, leaving the template's own
    /// formatting, formulas, charts and layout intact.
    /// <para>
    /// This is the cheap route to formatted reports: the design stays in Excel
    /// where it is easy to author, and Grasshopper only supplies values. It
    /// avoids having to expose a full cell-styling API.
    /// </para>
    /// <para>
    /// Templates use MiniExcel's placeholder syntax:
    /// <list type="bullet">
    ///   <item><c>{{title}}</c> — a single scalar value.</item>
    ///   <item><c>{{rows.name}}</c>, <c>{{rows.qty}}</c> — a repeating collection
    ///         called <c>rows</c>; the template row is duplicated once per item.</item>
    /// </list>
    /// </para>
    /// </summary>
    internal static class ExcelTemplate
    {
        /// <summary>
        /// Builds the value bag MiniExcel expects from parallel key/column lists.
        /// <para>
        /// A key without a dot becomes a scalar (the first value of its column).
        /// A dotted key <c>collection.field</c> contributes one field to a
        /// repeating collection, so several columns can share a collection name.
        /// </para>
        /// </summary>
        public static Dictionary<string, object> BuildValues(
            IReadOnlyList<string> keys, IReadOnlyList<IReadOnlyList<object>> columns)
        {
            if (keys is null) throw new ArgumentNullException(nameof(keys));
            if (columns is null) throw new ArgumentNullException(nameof(columns));

            var values = new Dictionary<string, object>(StringComparer.Ordinal);

            // collection name -> (field name -> column values)
            var collections = new Dictionary<string, List<(string Field, IReadOnlyList<object> Values)>>(
                StringComparer.Ordinal);
            var order = new List<string>();

            for (var i = 0; i < keys.Count; i++)
            {
                var key = keys[i]?.Trim();
                if (string.IsNullOrEmpty(key)) continue;

                var column = i < columns.Count ? columns[i] : Array.Empty<object>();

                var dot = key.IndexOf('.');
                if (dot <= 0 || dot == key.Length - 1)
                {
                    // Scalar placeholder: {{key}}
                    values[key] = column.Count > 0 ? column[0] : null;
                    continue;
                }

                var collectionName = key.Substring(0, dot);
                var field = key.Substring(dot + 1);

                if (!collections.TryGetValue(collectionName, out var fields))
                {
                    fields = new List<(string, IReadOnlyList<object>)>();
                    collections[collectionName] = fields;
                    order.Add(collectionName);
                }
                fields.Add((field, column));
            }

            // Pivot each collection's columns into a list of row dictionaries,
            // which is the shape MiniExcel's repeating-row templating expects.
            foreach (var name in order)
            {
                var fields = collections[name];
                var rowCount = fields.Count == 0 ? 0 : fields.Max(f => f.Values.Count);

                var rows = new List<IDictionary<string, object>>(rowCount);
                for (var r = 0; r < rowCount; r++)
                {
                    var row = new Dictionary<string, object>(StringComparer.Ordinal);
                    foreach (var (field, column) in fields)
                        row[field] = r < column.Count ? column[r] : null;
                    rows.Add(row);
                }

                values[name] = rows;
            }

            return values;
        }

        /// <summary>
        /// Renders <paramref name="templatePath"/> into <paramref name="outputPath"/>
        /// using the supplied values.
        /// </summary>
        public static void Fill(
            string templatePath, string outputPath, Dictionary<string, object> values, bool overwrite)
        {
            if (string.IsNullOrWhiteSpace(templatePath))
                throw new ArgumentException("Template path is required.", nameof(templatePath));
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Output path is required.", nameof(outputPath));

            var template = Path.GetFullPath(templatePath);
            var output = Path.GetFullPath(outputPath);

            if (!File.Exists(template))
                throw new FileNotFoundException("Template not found: " + template, template);

            // Rendering a template onto itself would consume the placeholders
            // and destroy the template.
            if (string.Equals(template, output, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "Output path must differ from the template path, otherwise the template would be overwritten and its placeholders lost.");

            if (File.Exists(output) && !overwrite)
                throw new InvalidOperationException(
                    $"'{Path.GetFileName(output)}' already exists and Overwrite? is false.");

            var directory = Path.GetDirectoryName(output);
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

            var config = new OpenXmlConfiguration
            {
                // A placeholder with no matching key should blank out rather
                // than abort the whole render.
                IgnoreTemplateParameterMissing = true,
            };

            MiniExcelLib.MiniExcel.Templaters
                .GetOpenXmlTemplater()
                .FillTemplate(output, template, values, overwrite, config);

            ExcelReadCache.Invalidate(output);
        }
    }
}

