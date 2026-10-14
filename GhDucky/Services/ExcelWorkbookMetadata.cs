using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using GhDucky.Utils;

namespace GhDucky.Services
{
    /// <summary>A worksheet as listed in the workbook part.</summary>
    internal sealed class ExcelSheetInfo
    {
        public string Name { get; init; }
        public int Index { get; init; }
        public bool Hidden { get; init; }
        public string UsedRange { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
    }

    /// <summary>The extent of the cells a worksheet actually uses.</summary>
    internal sealed class ExcelSheetDimension
    {
        public static readonly ExcelSheetDimension Empty =
            new() { UsedRange = string.Empty, RowCount = 0, ColumnCount = 0 };

        /// <summary>The used range, e.g. <c>B2:D11</c>, or empty for a blank sheet.</summary>
        public string UsedRange { get; init; }

        public int RowCount { get; init; }
        public int ColumnCount { get; init; }
    }

    /// <summary>A defined name (named range) and the reference it points at.</summary>
    internal sealed class ExcelDefinedName
    {
        public string Name { get; init; }

        /// <summary>The raw reference, e.g. <c>Data!$B$2:$C$3</c>.</summary>
        public string Reference { get; init; }

        /// <summary>Null for workbook-scoped names, otherwise the owning sheet.</summary>
        public string Scope { get; init; }
    }

    /// <summary>An Excel Table (ListObject) and its anchoring range.</summary>
    internal sealed class ExcelTableInfo
    {
        public string Name { get; init; }
        public string Sheet { get; init; }

        /// <summary>The table extent, e.g. <c>A1:C10</c>.</summary>
        public string Reference { get; init; }

        public IReadOnlyList<string> Columns { get; init; }
    }

    /// <summary>
    /// Reads workbook structure straight out of the xlsx package.
    /// <para>
    /// MiniExcel surfaces sheet names, sheet state and dimensions, but does not
    /// expose <c>definedNames</c> (named ranges) or enumerate Tables. Rather
    /// than take a second heavyweight dependency for that, the handful of parts
    /// involved are read directly — an xlsx is just a zip of XML, and these
    /// parts are small and stable.
    /// </para>
    /// </summary>
    internal static class ExcelWorkbookMetadata
    {
        private static readonly XNamespace Main =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private static readonly XNamespace Rel =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private static readonly XNamespace PkgRel =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        private static readonly XmlReaderSettings SheetReaderSettings = new()
        {
            IgnoreWhitespace = true,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
        };

        /// <summary>
        /// Measures every sheet's used range, keyed by sheet name.
        /// <para>
        /// MiniExcel exposes dimensions too, but its 2.0 preview throws
        /// IndexOutOfRangeException whenever a sheet's <c>dimension</c> element
        /// names a single cell (<c>ref="A1"</c>) rather than a range — which is
        /// what Excel writes for an empty or one-cell sheet.
        /// </para>
        /// </summary>
        public static IReadOnlyDictionary<string, ExcelSheetDimension> ReadSheetDimensions(string path)
        {
            var result = new Dictionary<string, ExcelSheetDimension>(StringComparer.OrdinalIgnoreCase);

            // Shared access so a workbook open in Excel can still be inspected.
            using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var zip = new ZipArchive(file, ZipArchiveMode.Read);

            var workbook = LoadXml(zip, "xl/workbook.xml");
            if (workbook is null) return result;

            var workbookRels = ReadRelationships(zip, "xl/_rels/workbook.xml.rels", "xl");

            foreach (var sheet in workbook.Root?.Element(Main + "sheets")?.Elements(Main + "sheet")
                                  ?? Enumerable.Empty<XElement>())
            {
                var name = (string)sheet.Attribute("name");
                if (string.IsNullOrEmpty(name)) continue;

                var rid = (string)sheet.Attribute(Rel + "id");
                result[name] = rid != null && workbookRels.TryGetValue(rid, out var part)
                    ? ReadSheetDimension(zip, part)
                    : ExcelSheetDimension.Empty;
            }

            return result;
        }

        private static ExcelSheetDimension ReadSheetDimension(ZipArchive zip, string partPath)
        {
            var entry = FindEntry(zip, partPath);
            if (entry is null) return ExcelSheetDimension.Empty;

            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, SheetReaderSettings);

            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element) continue;

                // A range <dimension> is authoritative; a single-cell one is
                // ambiguous (Excel uses A1 for an empty sheet) so the rows decide.
                if (reader.LocalName == "dimension")
                {
                    if (TryParseExtent(reader.GetAttribute("ref"), out var declared)) return declared;
                }
                else if (reader.LocalName == "sheetData")
                {
                    return ScanSheetData(reader);
                }
            }

            return ExcelSheetDimension.Empty;
        }

        private static bool TryParseExtent(string reference, out ExcelSheetDimension dimension)
        {
            dimension = null;
            if (string.IsNullOrWhiteSpace(reference)) return false;

            var parts = reference.Split(':');
            if (parts.Length != 2) return false;

            if (!ExcelAddress.TryParseCell(parts[0], out var r1, out var c1)) return false;
            if (!ExcelAddress.TryParseCell(parts[1], out var r2, out var c2)) return false;

            dimension = Build(Math.Min(r1, r2), Math.Min(c1, c2), Math.Max(r1, r2), Math.Max(c1, c2));
            return true;
        }

        /// <summary>Derives the used range from the cells present in sheetData.</summary>
        private static ExcelSheetDimension ScanSheetData(XmlReader reader)
        {
            var firstRow = int.MaxValue;
            var firstColumn = int.MaxValue;
            var lastRow = 0;
            var lastColumn = 0;

            var row = 0;
            var column = 0;

            using var sheetData = reader.ReadSubtree();
            while (sheetData.Read())
            {
                if (sheetData.NodeType != XmlNodeType.Element) continue;

                if (sheetData.LocalName == "row")
                {
                    row = ParseIndex(sheetData.GetAttribute("r"), row + 1);
                    column = 0;
                    continue;
                }

                if (sheetData.LocalName != "c") continue;

                // Both row and cell references are optional; position then decides.
                var cellRow = row;
                if (ExcelAddress.TryParseCell(sheetData.GetAttribute("r"), out var r, out var c))
                {
                    cellRow = r;
                    column = c;
                }
                else
                {
                    column++;
                }

                if (cellRow < firstRow) firstRow = cellRow;
                if (cellRow > lastRow) lastRow = cellRow;
                if (column < firstColumn) firstColumn = column;
                if (column > lastColumn) lastColumn = column;
            }

            return lastRow == 0 || lastColumn == 0
                ? ExcelSheetDimension.Empty
                : Build(firstRow, firstColumn, lastRow, lastColumn);
        }

        private static ExcelSheetDimension Build(int startRow, int startColumn, int endRow, int endColumn) =>
            new()
            {
                UsedRange = $"{ExcelAddress.Cell(startRow, startColumn)}:{ExcelAddress.Cell(endRow, endColumn)}",
                RowCount = endRow - startRow + 1,
                ColumnCount = endColumn - startColumn + 1,
            };

        private static int ParseIndex(string value, int fallback) =>
            int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) && index > 0
                ? index
                : fallback;

        /// <summary>
        /// Reads every defined name in the workbook, skipping Excel's internal
        /// built-ins (<c>_xlnm.Print_Area</c> and friends).
        /// </summary>
        public static IReadOnlyList<ExcelDefinedName> ReadDefinedNames(string path)
        {
            using var zip = ZipFile.OpenRead(path);
            return ReadDefinedNames(zip);
        }

        private static IReadOnlyList<ExcelDefinedName> ReadDefinedNames(ZipArchive zip)
        {
            var result = new List<ExcelDefinedName>();

            var workbook = LoadXml(zip, "xl/workbook.xml");
            if (workbook is null) return result;

            // localSheetId indexes into the <sheets> order, so build that first.
            var sheetNames = workbook.Root?
                .Element(Main + "sheets")?
                .Elements(Main + "sheet")
                .Select(e => (string)e.Attribute("name"))
                .ToList() ?? new List<string>();

            var defined = workbook.Root?.Element(Main + "definedNames");
            if (defined is null) return result;

            foreach (var dn in defined.Elements(Main + "definedName"))
            {
                var name = (string)dn.Attribute("name");
                if (string.IsNullOrWhiteSpace(name)) continue;

                // Built-in names (print areas, filter databases, ...) are noise.
                if (name.StartsWith("_xlnm", StringComparison.OrdinalIgnoreCase)) continue;

                string scope = null;
                var localSheetId = (string)dn.Attribute("localSheetId");
                if (localSheetId != null &&
                    int.TryParse(localSheetId, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx) &&
                    idx >= 0 && idx < sheetNames.Count)
                {
                    scope = sheetNames[idx];
                }

                result.Add(new ExcelDefinedName
                {
                    Name = name,
                    Reference = (dn.Value ?? string.Empty).Trim(),
                    Scope = scope,
                });
            }

            return result;
        }

        /// <summary>
        /// Resolves a defined name to its reference. Sheet-scoped names are
        /// preferred when <paramref name="preferredScope"/> matches; otherwise
        /// the first workbook-scoped match wins.
        /// </summary>
        public static string ResolveName(string path, string name, string preferredScope = null)
        {
            var names = ReadDefinedNames(path);

            var scoped = names.FirstOrDefault(n =>
                string.Equals(n.Name, name, StringComparison.OrdinalIgnoreCase) &&
                preferredScope != null &&
                string.Equals(n.Scope, preferredScope, StringComparison.OrdinalIgnoreCase));
            if (scoped != null) return scoped.Reference;

            var global = names.FirstOrDefault(n =>
                string.Equals(n.Name, name, StringComparison.OrdinalIgnoreCase) && n.Scope is null);
            if (global != null) return global.Reference;

            var any = names.FirstOrDefault(n =>
                string.Equals(n.Name, name, StringComparison.OrdinalIgnoreCase));
            if (any != null) return any.Reference;

            // Fall back to Tables, which share the same "addressable by name" idea.
            var table = ReadTables(path)
                .FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase));

            return table is null
                ? null
                : (string.IsNullOrEmpty(table.Sheet) ? table.Reference : table.Sheet + "!" + table.Reference);
        }

        /// <summary>Enumerates every Excel Table, resolved back to its owning sheet.</summary>
        public static IReadOnlyList<ExcelTableInfo> ReadTables(string path)
        {
            using var zip = ZipFile.OpenRead(path);
            var result = new List<ExcelTableInfo>();

            var workbook = LoadXml(zip, "xl/workbook.xml");
            if (workbook is null) return result;

            // rId -> worksheet part path
            var workbookRels = ReadRelationships(zip, "xl/_rels/workbook.xml.rels", "xl");

            foreach (var sheet in workbook.Root?.Element(Main + "sheets")?.Elements(Main + "sheet")
                                  ?? Enumerable.Empty<XElement>())
            {
                var sheetName = (string)sheet.Attribute("name");
                var rid = (string)sheet.Attribute(Rel + "id");
                if (rid is null || !workbookRels.TryGetValue(rid, out var sheetPart)) continue;

                // Each worksheet's own rels point at its table parts.
                var dir = GetDirectory(sheetPart);
                var relsPath = $"{dir}/_rels/{Path.GetFileName(sheetPart)}.rels";
                var sheetRels = ReadRelationships(zip, relsPath, dir);

                foreach (var target in sheetRels.Values)
                {
                    if (target.IndexOf("/tables/", StringComparison.OrdinalIgnoreCase) < 0) continue;

                    var tableXml = LoadXml(zip, target);
                    var root = tableXml?.Root;
                    if (root is null) continue;

                    var columns = root.Element(Main + "tableColumns")?
                        .Elements(Main + "tableColumn")
                        .Select(c => (string)c.Attribute("name"))
                        .ToList() ?? new List<string>();

                    result.Add(new ExcelTableInfo
                    {
                        Name = (string)root.Attribute("displayName") ?? (string)root.Attribute("name"),
                        Sheet = sheetName,
                        Reference = (string)root.Attribute("ref"),
                        Columns = columns,
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Reads the relationship file at <paramref name="relsPath"/> and returns
        /// rId → package-relative target path.
        /// </summary>
        private static Dictionary<string, string> ReadRelationships(
            ZipArchive zip, string relsPath, string baseDir)
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            var doc = LoadXml(zip, relsPath);
            if (doc?.Root is null) return map;

            foreach (var r in doc.Root.Elements(PkgRel + "Relationship"))
            {
                var id = (string)r.Attribute("Id");
                var target = (string)r.Attribute("Target");
                if (id is null || target is null) continue;

                map[id] = NormalisePath(baseDir, target);
            }
            return map;
        }

        /// <summary>Resolves a possibly relative relationship target to a zip entry path.</summary>
        private static string NormalisePath(string baseDir, string target)
        {
            target = target.Replace('\\', '/');
            if (target.StartsWith("/", StringComparison.Ordinal))
                return target.TrimStart('/');

            var segments = new List<string>();
            if (!string.IsNullOrEmpty(baseDir))
                segments.AddRange(baseDir.Split('/'));

            foreach (var part in target.Split('/'))
            {
                if (part == "." || part.Length == 0) continue;
                if (part == "..")
                {
                    if (segments.Count > 0) segments.RemoveAt(segments.Count - 1);
                    continue;
                }
                segments.Add(part);
            }
            return string.Join("/", segments);
        }

        private static string GetDirectory(string entryPath)
        {
            var i = entryPath.LastIndexOf('/');
            return i < 0 ? string.Empty : entryPath.Substring(0, i);
        }

        private static ZipArchiveEntry FindEntry(ZipArchive zip, string entryPath) =>
            zip.GetEntry(entryPath)
            ?? zip.Entries.FirstOrDefault(e =>
                string.Equals(e.FullName, entryPath, StringComparison.OrdinalIgnoreCase));

        private static XDocument LoadXml(ZipArchive zip, string entryPath)
        {
            var entry = FindEntry(zip, entryPath);
            if (entry is null) return null;

            using var stream = entry.Open();
            try
            {
                return XDocument.Load(stream);
            }
            catch
            {
                // A malformed or unexpected part should degrade gracefully rather
                // than take the whole Info component down.
                return null;
            }
        }
    }
}

