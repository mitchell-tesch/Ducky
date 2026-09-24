using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using GhDucky.Utils;
using MiniExcelLib.OpenXml;

namespace GhDucky.Services
{
    /// <summary>Outcome of a range read.</summary>
    internal sealed class ExcelReadResult
    {
        /// <summary>Row-major grid of CLR values; nulls represent empty cells.</summary>
        public object[][] Grid { get; init; }

        /// <summary>The address actually read, e.g. <c>Data!B2:D11</c>.</summary>
        public string Address { get; init; }

        /// <summary>
        /// Number of cells in the range that contain a formula but no cached
        /// result. These read as null and stay null until Excel next opens and
        /// recalculates the workbook — notably true of formulas Ducky itself
        /// wrote.
        /// </summary>
        public int UncachedFormulaCount { get; init; }

        /// <summary>Address of the first uncached formula cell, for diagnostics.</summary>
        public string FirstUncachedFormula { get; init; }

        public int RowCount => Grid?.Length ?? 0;
        public int ColumnCount => Grid is { Length: > 0 } ? Grid[0].Length : 0;
    }

    /// <summary>
    /// Reads cell ranges out of xlsx/xlsm workbooks.
    /// <para>
    /// Backed by MiniExcel's streaming OpenXML reader, which handles A1 ranges,
    /// open-ended ranges and native cell typing (numbers as <see cref="double"/>,
    /// dates as <see cref="DateTime"/>, booleans as <see cref="bool"/>, and
    /// formula cells as their cached value).
    /// </para>
    /// </summary>
    internal static class ExcelReader
    {
        /// <summary>
        /// Reads a range. When <paramref name="expand"/> is true and the address
        /// is only an anchor cell, the extent grows to the enclosing contiguous
        /// block — a fully blank row bounds it vertically and a fully blank
        /// column bounds it horizontally (Excel's "current region").
        /// </summary>
        /// <param name="path">Path to the .xlsx/.xlsm workbook.</param>
        /// <param name="address">The cell or range to read.</param>
        /// <param name="expand">
        /// Expand an anchor-only address to its contiguous region. Ignored when
        /// the address carries an explicit extent.
        /// </param>
        /// <param name="formulaText">
        /// When true, cells containing a formula yield the formula text
        /// (prefixed with <c>=</c>) instead of the cached result.
        /// </param>
        public static ExcelReadResult Read(string path, ExcelAddress address, bool expand, bool formulaText)
        {
            if (address is null) throw new ArgumentNullException(nameof(address));

            var full = Path.GetFullPath(path);

            var importer = MiniExcelLib.MiniExcel.Importers.GetOpenXmlImporter();

            // Keep blank rows so sheet row positions stay aligned with grid rows.
            var config = new OpenXmlConfiguration
            {
                IgnoreEmptyRows = false,
                FillMergedCells = true,
            };

            var resolvedSheet = ResolveSheetName(importer, full, address.Sheet);

            var startRow = address.StartRow;
            var startCol = address.StartColumn;
            int? endRow = address.EndRow;
            int? endCol = address.EndColumn;

            // An open-ended anchor with expansion disabled means "just this cell".
            if (address.IsAnchorOnly && !expand)
            {
                endRow = startRow;
                endCol = startCol;
            }

            List<IDictionary<string, object>> rawRows;
            using (var stream = OpenShared(full))
            {
                rawRows = importer
                    .QueryRange(stream, false, resolvedSheet, startRow, startCol, endRow, endCol, config, true)
                    .Cast<IDictionary<string, object>>()
                    .ToList();
            }

            var grid = BuildGrid(rawRows, startCol, endCol);

            // One streaming pass over the sheet part handles both formula jobs:
            //  - formulaText: replace cached results with the formula text;
            //  - otherwise:   count formulas whose cached result is missing, so
            //                 the component can warn instead of silently
            //                 yielding nulls.
            // Formulas must be handled *before* trimming, because a formula
            // Ducky wrote has no cached <v> and so reads as null; trimming first
            // would discard the very column we need to inspect.
            var uncachedCount = 0;
            string firstUncached = null;
            bool[][] occupied = null;

            if (grid.Length > 0 && (formulaText || HasNulls(grid)))
            {
                using var stream = OpenShared(full);
                (uncachedCount, firstUncached, occupied) =
                    ScanFormulas(stream, resolvedSheet, startRow, startCol, grid, formulaText);
            }

            grid = address.IsAnchorOnly && expand
                ? TrimToContiguousRegion(grid, occupied)
                : TrimTrailingEmpty(grid, occupied);

            var rowCount = grid.Length;
            var colCount = rowCount > 0 ? grid[0].Length : 0;

            return new ExcelReadResult
            {
                Grid = grid,
                UncachedFormulaCount = uncachedCount,
                FirstUncachedFormula = firstUncached,
                Address = rowCount == 0 || colCount == 0
                    ? ExcelAddress.Format(resolvedSheet, startRow, startCol, startRow, startCol)
                    : ExcelAddress.Format(resolvedSheet, startRow, startCol,
                        startRow + rowCount - 1, startCol + colCount - 1),
            };
        }

        /// <summary>True when any cell in the grid is empty.</summary>
        private static bool HasNulls(object[][] grid)
        {
            foreach (var row in grid)
            {
                foreach (var cell in row)
                {
                    if (cell is null) return true;
                }
            }
            return false;
        }

        /// <summary>Lists sheet names in workbook order.</summary>
        public static List<string> GetSheetNames(string path)
        {
            using var stream = OpenShared(Path.GetFullPath(path));
            return MiniExcelLib.MiniExcel.Importers.GetOpenXmlImporter().GetSheetNames(stream, true);
        }

        /// <summary>Lists sheets with their hidden state and used-range dimensions.</summary>
        public static List<ExcelSheetInfo> GetSheets(string path)
        {
            var full = Path.GetFullPath(path);
            var importer = MiniExcelLib.MiniExcel.Importers.GetOpenXmlImporter();

            List<MiniExcelLib.OpenXml.Models.SheetInfo> infos;
            using (var stream = OpenShared(full))
                infos = importer.GetSheetInformations(stream, true);

            // Dimensions are read in-house: MiniExcel's GetSheetDimensions throws
            // IndexOutOfRangeException on a single-cell <dimension ref="A1">, which
            // is exactly what Excel writes for an empty sheet.
            var dims = ExcelWorkbookMetadata.ReadSheetDimensions(full);

            var result = new List<ExcelSheetInfo>(infos.Count);
            foreach (var info in infos)
            {
                ExcelSheetDimension dim = null;
                if (info.Name != null) dims.TryGetValue(info.Name, out dim);

                result.Add(new ExcelSheetInfo
                {
                    Name = info.Name,
                    Index = (int)info.Index,
                    Hidden = info.State != MiniExcelLib.OpenXml.Models.SheetState.Visible,
                    UsedRange = dim?.UsedRange ?? string.Empty,
                    RowCount = dim?.RowCount ?? 0,
                    ColumnCount = dim?.ColumnCount ?? 0,
                });
            }
            return result;
        }

        /// <summary>
        /// Matches a requested sheet name case-insensitively, so "data" finds
        /// "Data". Returns null (meaning "first sheet") when none was requested.
        /// </summary>
        private static string ResolveSheetName(OpenXmlImporter importer, string path, string requested)
        {
            List<string> names;
            using (var stream = OpenShared(path))
                names = importer.GetSheetNames(stream, true);

            if (string.IsNullOrWhiteSpace(requested))
                return names.Count > 0 ? names[0] : null;

            var exact = names.FirstOrDefault(n => string.Equals(n, requested, StringComparison.Ordinal));
            if (exact != null) return exact;

            var loose = names.FirstOrDefault(n => string.Equals(n, requested, StringComparison.OrdinalIgnoreCase));
            if (loose != null) return loose;

            throw new ArgumentException(
                $"Sheet '{requested}' not found. Available: {string.Join(", ", names)}");
        }

        /// <summary>
        /// Converts MiniExcel's column-letter-keyed rows into a dense row-major grid.
        /// </summary>
        private static object[][] BuildGrid(
            List<IDictionary<string, object>> rows, int startCol, int? endCol)
        {
            if (rows.Count == 0) return Array.Empty<object[]>();

            // Work out the widest column actually returned.
            var maxCol = endCol ?? 0;
            if (endCol is null)
            {
                foreach (var row in rows)
                {
                    foreach (var key in row.Keys)
                    {
                        if (!TryColumnIndex(key, out var idx)) continue;
                        if (idx > maxCol) maxCol = idx;
                    }
                }
            }

            if (maxCol < startCol) return Array.Empty<object[]>();

            var width = maxCol - startCol + 1;
            var grid = new object[rows.Count][];

            for (var r = 0; r < rows.Count; r++)
            {
                var cells = new object[width];
                foreach (var kv in rows[r])
                {
                    if (!TryColumnIndex(kv.Key, out var idx)) continue;

                    var offset = idx - startCol;
                    if (offset < 0 || offset >= width) continue;

                    cells[offset] = Normalise(kv.Value);
                }
                grid[r] = cells;
            }

            return grid;
        }

        /// <summary>
        /// Expands to the enclosing contiguous block: the first entirely blank
        /// row terminates the block vertically, the first entirely blank column
        /// terminates it horizontally.
        /// </summary>
        private static object[][] TrimToContiguousRegion(object[][] grid, bool[][] occupied)
        {
            if (grid.Length == 0) return grid;

            var lastRow = grid.Length;
            for (var r = 0; r < grid.Length; r++)
            {
                if (IsRowEmpty(grid, occupied, r)) { lastRow = r; break; }
            }
            if (lastRow == 0) return Array.Empty<object[]>();

            var width = grid[0].Length;
            var lastCol = width;
            for (var c = 0; c < width; c++)
            {
                var empty = true;
                for (var r = 0; r < lastRow; r++)
                {
                    if (IsOccupied(grid, occupied, r, c)) { empty = false; break; }
                }
                if (empty) { lastCol = c; break; }
            }
            if (lastCol == 0) return Array.Empty<object[]>();

            return Crop(grid, lastRow, lastCol);
        }

        /// <summary>
        /// Drops trailing blank rows/columns from an explicitly-bounded read so
        /// an over-wide range does not produce a cloud of empty branches.
        /// </summary>
        private static object[][] TrimTrailingEmpty(object[][] grid, bool[][] occupied)
        {
            if (grid.Length == 0) return grid;

            var lastRow = grid.Length;
            while (lastRow > 0 && IsRowEmpty(grid, occupied, lastRow - 1)) lastRow--;
            if (lastRow == 0) return Array.Empty<object[]>();

            var width = grid[0].Length;
            var lastCol = width;
            while (lastCol > 0)
            {
                var empty = true;
                for (var r = 0; r < lastRow; r++)
                {
                    if (IsOccupied(grid, occupied, r, lastCol - 1)) { empty = false; break; }
                }
                if (!empty) break;
                lastCol--;
            }
            if (lastCol == 0) return Array.Empty<object[]>();

            if (lastRow == grid.Length && lastCol == width) return grid;

            return Crop(grid, lastRow, lastCol);
        }

        private static object[][] Crop(object[][] grid, int rows, int columns)
        {
            var result = new object[rows][];
            for (var r = 0; r < rows; r++)
            {
                var row = new object[columns];
                Array.Copy(grid[r], row, columns);
                result[r] = row;
            }
            return result;
        }

        /// <summary>
        /// A cell counts as occupied when it holds a value, or when the
        /// occupancy mask marks it (an uncached formula cell reads as null but
        /// is certainly not empty).
        /// </summary>
        private static bool IsOccupied(object[][] grid, bool[][] occupied, int row, int column) =>
            grid[row][column] != null || (occupied != null && occupied[row][column]);

        private static bool IsRowEmpty(object[][] grid, bool[][] occupied, int row)
        {
            var width = grid[row].Length;
            for (var c = 0; c < width; c++)
            {
                if (IsOccupied(grid, occupied, row, c)) return false;
            }
            return true;
        }

        /// <summary>
        /// Streams the sheet part looking at formula cells inside the read range.
        /// <para>
        /// When <paramref name="overlay"/> is true the cached result in the grid
        /// is replaced with the formula text. Otherwise the method counts
        /// formula cells whose cached result is absent — MiniExcel surfaces only
        /// the cached <c>&lt;v&gt;</c>, so such cells read as null.
        /// </para>
        /// <para>
        /// In the counting case an occupancy mask is also produced. A formula
        /// cell is <i>not</i> an empty cell, so it must not be trimmed away even
        /// though its value is null; the mask lets the trimming step know that.
        /// It is allocated only when such a cell is actually found, so the
        /// common path stays allocation-free.
        /// </para>
        /// </summary>
        private static (int Uncached, string First, bool[][] Occupied) ScanFormulas(
            Stream stream, string sheetName, int startRow, int startCol, object[][] grid, bool overlay)
        {
            var rowCount = grid.Length;
            var colCount = grid[0].Length;
            var endRow = startRow + rowCount - 1;
            var endCol = startCol + colCount - 1;

            var uncached = 0;
            string first = null;
            bool[][] occupied = null;

            using var doc = SpreadsheetDocument.Open(stream, false);
            var wbPart = doc.WorkbookPart;
            if (wbPart is null) return (0, null, null);

            var sheet = wbPart.Workbook.Descendants<Sheet>()
                            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
                        ?? wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
            if (sheet?.Id?.Value is null) return (0, null, null);

            if (wbPart.GetPartById(sheet.Id.Value) is not WorksheetPart wsPart) return (0, null, null);

            using var reader = OpenXmlReader.Create(wsPart);
            while (reader.Read())
            {
                if (reader.ElementType != typeof(Cell) || !reader.IsStartElement) continue;

                if (reader.LoadCurrentElement() is not Cell cell) continue;

                var formula = cell.CellFormula?.Text;
                if (string.IsNullOrEmpty(formula)) continue;

                var reference = cell.CellReference?.Value;
                if (reference is null) continue;
                if (!ExcelAddress.TryParseCell(reference, out var r, out var c)) continue;

                if (r < startRow || r > endRow || c < startCol || c > endCol) continue;

                if (overlay)
                {
                    grid[r - startRow][c - startCol] = "=" + formula;
                    continue;
                }

                // Not overlaying: flag formulas Excel has never evaluated.
                if (grid[r - startRow][c - startCol] is null)
                {
                    uncached++;
                    first ??= ExcelAddress.Format(sheetName, r, c, r, c);

                    occupied ??= BuildOccupancy(grid);
                    occupied[r - startRow][c - startCol] = true;
                }
            }

            return (uncached, first, occupied);
        }

        /// <summary>Seeds an occupancy mask from the grid's non-null cells.</summary>
        private static bool[][] BuildOccupancy(object[][] grid)
        {
            var mask = new bool[grid.Length][];
            for (var r = 0; r < grid.Length; r++)
            {
                var row = new bool[grid[r].Length];
                for (var c = 0; c < row.Length; c++)
                    row[c] = grid[r][c] != null;
                mask[r] = row;
            }
            return mask;
        }

        /// <summary>Normalises MiniExcel output into the CLR types Ducky maps to goo.</summary>
        private static object Normalise(object value)
        {
            if (value is null || value is DBNull) return null;

            // MiniExcel can surface an empty string for a styled-but-empty cell;
            // treat that as a genuine gap so data trees stay clean.
            if (value is string s && s.Length == 0) return null;

            return value;
        }

        private static bool TryColumnIndex(string key, out int index)
        {
            index = 0;
            if (string.IsNullOrEmpty(key)) return false;

            foreach (var ch in key)
            {
                var upper = char.ToUpperInvariant(ch);
                if (upper is < 'A' or > 'Z') return false;
                index = index * 26 + (upper - 'A' + 1);
            }
            return index > 0;
        }

        /// <summary>
        /// Opens a workbook for reading while allowing other processes (notably
        /// Excel) to keep it open.
        /// </summary>
        private static FileStream OpenShared(string path) =>
            new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    }
}







