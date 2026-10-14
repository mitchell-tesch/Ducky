using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using GhDucky.Utils;

namespace GhDucky.Services
{
    /// <summary>How a write interacts with an existing workbook.</summary>
    internal enum ExcelWriteMode
    {
        /// <summary>
        /// Modify only the target cells. Every other cell, its formatting, the
        /// sheet's formulas and all other sheets are left untouched.
        /// </summary>
        InPlace = 0,

        /// <summary>Clear the target sheet first, then write. Other sheets survive.</summary>
        OverwriteSheet = 1,

        /// <summary>Discard the file and build a fresh single-sheet workbook.</summary>
        NewWorkbook = 2,
    }

    /// <summary>What a clear operation removes.</summary>
    internal enum ExcelClearMode
    {
        /// <summary>Remove cell values and formulas, keeping formatting.</summary>
        Values = 0,

        /// <summary>Remove the cells entirely, discarding their formatting too.</summary>
        ValuesAndFormatting = 1,

        /// <summary>Remove every row on the sheet. The Range input is ignored.</summary>
        EntireSheet = 2,
    }

    /// <summary>Outcome of a write.</summary>
    internal sealed class ExcelWriteResult
    {
        /// <summary>The address actually written, e.g. <c>Data!B2:D11</c>.</summary>
        public string Address { get; init; }

        /// <summary>Number of cells written, or cleared for a clear operation.</summary>
        public int CellsWritten { get; init; }

        /// <summary>True when the supplied data overflowed an explicit range and was clipped.</summary>
        public bool Clipped { get; init; }
    }

    /// <summary>
    /// Writes cell ranges into xlsx/xlsm workbooks using the OpenXML SDK.
    /// <para>
    /// The SDK is used here rather than MiniExcel because MiniExcel has no
    /// cell-level write: its only write verbs replace an entire sheet or an
    /// entire file, which destroys formatting, formulas and neighbouring data.
    /// </para>
    /// <para>
    /// Writes are staged into a sibling temporary file and then swapped in, so
    /// an interrupted write cannot leave a half-rewritten (corrupt) zip behind.
    /// </para>
    /// </summary>
    internal static class ExcelWriter
    {
        /// <summary>
        /// Writes <paramref name="grid"/> (row-major, in sheet orientation) with
        /// its top-left cell at the address anchor.
        /// <para>
        /// When the address carries an explicit extent, data beyond it is
        /// clipped. Null entries clear the target cell but keep its formatting.
        /// Strings beginning with <c>=</c> are written as formulas when
        /// <paramref name="allowFormulas"/> is set.
        /// </para>
        /// </summary>
        public static ExcelWriteResult Write(
            string path,
            ExcelAddress address,
            object[][] grid,
            ExcelWriteMode mode,
            bool allowFormulas)
        {
            if (address is null) throw new ArgumentNullException(nameof(address));
            grid ??= Array.Empty<object[]>();

            var full = Path.GetFullPath(path);
            var directory = Path.GetDirectoryName(full);
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

            var exists = File.Exists(full);
            var effectiveMode = exists ? mode : ExcelWriteMode.NewWorkbook;

            // ---- clip to an explicit extent ---------------------------------
            var clipped = false;
            var rowLimit = address.EndRow is null ? int.MaxValue : address.EndRow.Value - address.StartRow + 1;
            var colLimit = address.EndColumn is null ? int.MaxValue : address.EndColumn.Value - address.StartColumn + 1;

            if (grid.Length > rowLimit)
            {
                grid = grid.Take(rowLimit).ToArray();
                clipped = true;
            }

            var width = grid.Length == 0 ? 0 : grid.Max(r => r?.Length ?? 0);
            if (width > colLimit)
            {
                width = colLimit;
                clipped = true;
            }

            // Guard against running off the end of the sheet.
            if (address.StartRow + Math.Max(grid.Length, 1) - 1 > ExcelAddress.MaxRow ||
                address.StartColumn + Math.Max(width, 1) - 1 > ExcelAddress.MaxColumn)
            {
                throw new InvalidOperationException(
                    "The data does not fit on the worksheet from the given anchor.");
            }

            var sheetName = address.Sheet;
            var temp = Path.Combine(
                string.IsNullOrEmpty(directory) ? "." : directory,
                "." + Path.GetFileName(full) + "." + Guid.NewGuid().ToString("N") + ".tmp");

            var cellsWritten = 0;

            try
            {
                if (effectiveMode == ExcelWriteMode.NewWorkbook)
                    CreateEmptyWorkbook(temp, string.IsNullOrWhiteSpace(sheetName) ? "Sheet1" : sheetName,
                        DocumentTypeFor(full));
                else
                    File.Copy(full, temp, true);

                using (var doc = SpreadsheetDocument.Open(temp, true))
                {
                    var wbPart = doc.WorkbookPart
                                 ?? throw new InvalidOperationException("Workbook part is missing or the file is not a valid xlsx.");

                    var (wsPart, resolvedSheet) = ResolveOrCreateSheet(wbPart, sheetName);
                    sheetName = resolvedSheet;

                    var worksheet = wsPart.Worksheet;
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData is null)
                    {
                        sheetData = new SheetData();
                        worksheet.AppendChild(sheetData);
                    }

                    if (effectiveMode == ExcelWriteMode.OverwriteSheet)
                        sheetData.RemoveAllChildren<Row>();

                    var styles = new DateStyleProvider(wbPart);

                    for (var r = 0; r < grid.Length; r++)
                    {
                        var source = grid[r];
                        if (source is null) continue;

                        var rowIndex = (uint)(address.StartRow + r);
                        var row = GetOrCreateRow(sheetData, rowIndex);

                        var count = Math.Min(source.Length, width);
                        for (var c = 0; c < count; c++)
                        {
                            var columnIndex = address.StartColumn + c;
                            var cell = GetOrCreateCell(row, rowIndex, columnIndex);
                            SetCellValue(cell, source[c], allowFormulas, styles);
                            cellsWritten++;
                        }
                    }

                    UpdateDimension(worksheet, sheetData);

                    // Any write can feed an existing formula, so force Excel to
                    // recalculate on open and drop the now-stale calc chain.
                    ForceRecalculation(wbPart);

                    worksheet.Save();
                    wbPart.Workbook.Save();
                }

                Commit(temp, full, exists);
            }
            finally
            {
                TryDelete(temp);
            }

            ExcelReadCache.Invalidate(full);

            var rows = grid.Length;
            var cols = width;

            return new ExcelWriteResult
            {
                Address = rows == 0 || cols == 0
                    ? ExcelAddress.Format(sheetName, address.StartRow, address.StartColumn, address.StartRow, address.StartColumn)
                    : ExcelAddress.Format(sheetName, address.StartRow, address.StartColumn,
                        address.StartRow + rows - 1, address.StartColumn + cols - 1),
                CellsWritten = cellsWritten,
                Clipped = clipped,
            };
        }

        /// <summary>
        /// Clears cells without writing anything in their place.
        /// <para>
        /// An anchor-only address clears exactly that one cell. Clearing is
        /// destructive, so it deliberately does not auto-expand.
        /// </para>
        /// </summary>
        public static ExcelWriteResult Clear(string path, ExcelAddress address, ExcelClearMode mode)
        {
            if (address is null) throw new ArgumentNullException(nameof(address));

            var full = Path.GetFullPath(path);
            if (!File.Exists(full))
                throw new FileNotFoundException("Workbook not found: " + full, full);

            var directory = Path.GetDirectoryName(full);
            var temp = Path.Combine(
                string.IsNullOrEmpty(directory) ? "." : directory,
                "." + Path.GetFileName(full) + "." + Guid.NewGuid().ToString("N") + ".tmp");

            var startRow = address.StartRow;
            var startCol = address.StartColumn;
            var endRow = address.EndRow ?? startRow;
            var endCol = address.EndColumn ?? startCol;

            var sheetName = address.Sheet;
            var cleared = 0;

            try
            {
                File.Copy(full, temp, true);

                using (var doc = SpreadsheetDocument.Open(temp, true))
                {
                    var wbPart = doc.WorkbookPart
                                 ?? throw new InvalidOperationException("Workbook part is missing or the file is not a valid xlsx.");

                    var (wsPart, resolvedSheet) = ResolveExistingSheet(wbPart, sheetName);
                    sheetName = resolvedSheet;

                    var worksheet = wsPart.Worksheet;
                    var sheetData = worksheet.GetFirstChild<SheetData>();

                    if (sheetData != null)
                    {
                        if (mode == ExcelClearMode.EntireSheet)
                        {
                            cleared = sheetData.Elements<Row>().Sum(r => r.Elements<Cell>().Count());
                            sheetData.RemoveAllChildren<Row>();
                        }
                        else
                        {
                            cleared = ClearRange(sheetData, startRow, startCol, endRow, endCol,
                                mode == ExcelClearMode.ValuesAndFormatting);
                        }
                    }

                    UpdateDimension(worksheet, worksheet.GetFirstChild<SheetData>() ?? new SheetData());
                    ForceRecalculation(wbPart);

                    worksheet.Save();
                    wbPart.Workbook.Save();
                }

                Commit(temp, full, true);
            }
            finally
            {
                TryDelete(temp);
            }

            ExcelReadCache.Invalidate(full);

            return new ExcelWriteResult
            {
                Address = mode == ExcelClearMode.EntireSheet
                    ? (sheetName ?? string.Empty)
                    : ExcelAddress.Format(sheetName, startRow, startCol, endRow, endCol),
                CellsWritten = cleared,
                Clipped = false,
            };
        }

        /// <summary>
        /// Clears the cells inside the given bounds. Returns how many cells were
        /// actually affected (absent cells are not counted).
        /// </summary>
        private static int ClearRange(
            SheetData sheetData, int startRow, int startCol, int endRow, int endCol, bool removeFormatting)
        {
            var cleared = 0;

            foreach (var row in sheetData.Elements<Row>().ToList())
            {
                var rowIndex = (int)(row.RowIndex?.Value ?? 0U);
                if (rowIndex < startRow || rowIndex > endRow) continue;

                foreach (var cell in row.Elements<Cell>().ToList())
                {
                    var reference = cell.CellReference?.Value;
                    if (reference is null) continue;
                    if (!ExcelAddress.TryParseCell(reference, out _, out var col)) continue;
                    if (col < startCol || col > endCol) continue;

                    if (removeFormatting)
                    {
                        cell.Remove();
                    }
                    else
                    {
                        cell.CellFormula = null;
                        cell.CellValue = null;
                        cell.InlineString = null;
                        cell.DataType = null;
                    }
                    cleared++;
                }

                // Drop rows left with no cells so the file does not accumulate
                // empty <row> elements over repeated clears.
                if (!row.Elements<Cell>().Any()) row.Remove();
            }

            return cleared;
        }

        /// <summary>
        /// Swaps the staged file into place. <see cref="File.Replace(string,string,string)"/>
        /// is preferred because it is atomic and preserves the destination's
        /// attributes and ACLs.
        /// </summary>
        private static void Commit(string temp, string target, bool targetExisted)
        {
            if (!targetExisted)
            {
                File.Move(temp, target);
                return;
            }

            try
            {
                File.Replace(temp, target, null, true);
            }
            catch (PlatformNotSupportedException)
            {
                File.Copy(temp, target, true);
                TryDelete(temp);
            }
            catch (IOException)
            {
                // Replace can fail across some network shares; fall back to copy.
                File.Copy(temp, target, true);
                TryDelete(temp);
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
            }
            catch
            {
                // A leftover temp file is not worth failing the solve over.
            }
        }

        /// <summary>
        /// Chooses the OpenXML document type from the <i>destination</i>
        /// extension. A <c>.xlsm</c> file must be a MacroEnabledWorkbook,
        /// otherwise Excel reports a "file format and extension don't match"
        /// warning on open.
        /// <para>
        /// This deliberately takes the final path rather than the staging file,
        /// whose extension is always <c>.tmp</c>.
        /// </para>
        /// </summary>
        private static SpreadsheetDocumentType DocumentTypeFor(string destinationPath) =>
            Path.GetExtension(destinationPath).Equals(".xlsm", StringComparison.OrdinalIgnoreCase)
                ? SpreadsheetDocumentType.MacroEnabledWorkbook
                : SpreadsheetDocumentType.Workbook;

        /// <summary>Creates a minimal single-sheet workbook at the given path.</summary>
        private static void CreateEmptyWorkbook(string path, string sheetName, SpreadsheetDocumentType documentType)
        {
            ValidateSheetName(sheetName);

            using var doc = SpreadsheetDocument.Create(path, documentType);

            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();

            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData());

            var sheets = wbPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = 1U,
                Name = sheetName,
            });

            wbPart.Workbook.Save();
        }

        /// <summary>
        /// Finds the named sheet (case-insensitively) or appends a new one.
        /// A null/blank name selects the first sheet.
        /// </summary>
        private static (WorksheetPart Part, string Name) ResolveOrCreateSheet(WorkbookPart wbPart, string sheetName)
        {
            var sheets = wbPart.Workbook.GetFirstChild<Sheets>();
            if (sheets is null)
            {
                sheets = new Sheets();
                wbPart.Workbook.AppendChild(sheets);
            }

            var all = sheets.Elements<Sheet>().ToList();

            Sheet target;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                target = all.FirstOrDefault();
            }
            else
            {
                target = all.FirstOrDefault(s =>
                             string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal))
                         ?? all.FirstOrDefault(s =>
                             string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
            }

            if (target?.Id?.Value != null && wbPart.GetPartById(target.Id.Value) is WorksheetPart existing)
                return (existing, target.Name?.Value);

            // No match — add a new sheet rather than failing, so writing to a
            // brand-new tab is a one-input operation.
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData());

            var nextId = all.Count == 0 ? 1U : all.Max(s => s.SheetId?.Value ?? 0U) + 1U;
            var name = string.IsNullOrWhiteSpace(sheetName) ? "Sheet" + nextId : sheetName;
            ValidateSheetName(name);

            sheets.Append(new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = nextId,
                Name = name,
            });

            return (wsPart, name);
        }

        /// <summary>
        /// Finds an existing sheet, throwing a helpful error when it is absent.
        /// Used by operations (like clearing) where silently creating a sheet
        /// would be surprising.
        /// </summary>
        private static (WorksheetPart Part, string Name) ResolveExistingSheet(WorkbookPart wbPart, string sheetName)
        {
            var all = wbPart.Workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList() ?? new List<Sheet>();

            Sheet target;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                target = all.FirstOrDefault();
            }
            else
            {
                target = all.FirstOrDefault(s =>
                             string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal))
                         ?? all.FirstOrDefault(s =>
                             string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
            }

            if (target?.Id?.Value != null && wbPart.GetPartById(target.Id.Value) is WorksheetPart part)
                return (part, target.Name?.Value);

            var available = string.Join(", ", all.Select(s => s.Name?.Value));
            throw new ArgumentException(
                $"Sheet '{sheetName}' not found. Available: {available}");
        }

        /// <summary>
        /// Rejects sheet names Excel will not accept. Without this the file is
        /// written happily and then refused on open, which is a far harder
        /// failure to diagnose.
        /// </summary>
        private static void ValidateSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Sheet name cannot be empty.");

            if (name.Length > 31)
                throw new ArgumentException(
                    $"Sheet name '{name}' is {name.Length} characters; Excel allows at most 31.");

            var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
            var found = name.IndexOfAny(invalid);
            if (found >= 0)
                throw new ArgumentException(
                    $"Sheet name '{name}' contains '{name[found]}'. Excel forbids : \\ / ? * [ ] in sheet names.");

            if (name.StartsWith("'", StringComparison.Ordinal) || name.EndsWith("'", StringComparison.Ordinal))
                throw new ArgumentException($"Sheet name '{name}' cannot start or end with an apostrophe.");
        }

        /// <summary>Gets the row at <paramref name="rowIndex"/>, inserting it in order if absent.</summary>
        private static Row GetOrCreateRow(SheetData sheetData, uint rowIndex)
        {
            Row after = null;
            foreach (var row in sheetData.Elements<Row>())
            {
                var index = row.RowIndex?.Value ?? 0U;
                if (index == rowIndex) return row;
                if (index > rowIndex) break;
                after = row;
            }

            var created = new Row { RowIndex = rowIndex };
            if (after is null) sheetData.InsertAt(created, 0);
            else sheetData.InsertAfter(created, after);
            return created;
        }

        /// <summary>Gets the cell at the given column, inserting it in order if absent.</summary>
        private static Cell GetOrCreateCell(Row row, uint rowIndex, int columnIndex)
        {
            var reference = ExcelAddress.Cell((int)rowIndex, columnIndex);

            Cell after = null;
            foreach (var cell in row.Elements<Cell>())
            {
                var cellRef = cell.CellReference?.Value;
                if (cellRef is null) continue;

                if (string.Equals(cellRef, reference, StringComparison.OrdinalIgnoreCase))
                    return cell;

                if (!ExcelAddress.TryParseCell(cellRef, out _, out var col)) continue;
                if (col > columnIndex) break;
                after = cell;
            }

            var created = new Cell { CellReference = reference };
            if (after is null) row.InsertAt(created, 0);
            else row.InsertAfter(created, after);
            return created;
        }

        /// <summary>
        /// Writes a CLR value into a cell, preserving the cell's existing style
        /// (number format, fill, font, borders) wherever possible.
        /// </summary>
        private static void SetCellValue(Cell cell, object value, bool allowFormulas, DateStyleProvider styles)
        {
            // Clear the previous content but deliberately keep StyleIndex.
            cell.CellFormula = null;
            cell.CellValue = null;
            cell.InlineString = null;
            cell.DataType = null;

            if (value is null) return;

            switch (value)
            {
                case bool b:
                    cell.DataType = CellValues.Boolean;
                    cell.CellValue = new CellValue(b ? "1" : "0");
                    return;

                case DateTime dt:
                    cell.CellValue = new CellValue(dt.ToOADate().ToString("R", CultureInfo.InvariantCulture));
                    cell.DataType = CellValues.Number;
                    styles.ApplyDateFormat(cell, dt);
                    return;

                case DateTimeOffset dto:
                    cell.CellValue = new CellValue(dto.DateTime.ToOADate().ToString("R", CultureInfo.InvariantCulture));
                    cell.DataType = CellValues.Number;
                    styles.ApplyDateFormat(cell, dto.DateTime);
                    return;

                case double d:
                    WriteNumber(cell, d);
                    return;
                case float f:
                    WriteNumber(cell, f);
                    return;
                case decimal m:
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(m.ToString(CultureInfo.InvariantCulture));
                    return;
                case int i:
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(i.ToString(CultureInfo.InvariantCulture));
                    return;
                case long l:
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(l.ToString(CultureInfo.InvariantCulture));
                    return;
                case short s:
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(s.ToString(CultureInfo.InvariantCulture));
                    return;
                case byte by:
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(by.ToString(CultureInfo.InvariantCulture));
                    return;
            }

            var text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;

            if (text.Length == 0)
                return;

            if (allowFormulas && text.Length > 1 && text[0] == '=')
            {
                // No cached <v>: Excel recalculates on load (we also force a
                // full calculation below), which keeps the result honest.
                cell.CellFormula = new CellFormula(text.Substring(1));
                return;
            }

            // Inline strings avoid mutating the workbook-wide shared string
            // table, which would otherwise churn on every write.
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        }

        private static void WriteNumber(Cell cell, double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(value.ToString(CultureInfo.InvariantCulture)));
                return;
            }

            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(value.ToString("R", CultureInfo.InvariantCulture));
        }

        /// <summary>Recomputes the worksheet's declared dimension from its rows.</summary>
        private static void UpdateDimension(Worksheet worksheet, SheetData sheetData)
        {
            var minRow = int.MaxValue;
            var maxRow = 0;
            var minCol = int.MaxValue;
            var maxCol = 0;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var reference = cell.CellReference?.Value;
                    if (reference is null) continue;
                    if (!ExcelAddress.TryParseCell(reference, out var r, out var c)) continue;

                    if (r < minRow) minRow = r;
                    if (r > maxRow) maxRow = r;
                    if (c < minCol) minCol = c;
                    if (c > maxCol) maxCol = c;
                }
            }

            if (maxRow == 0 || maxCol == 0)
            {
                worksheet.SheetDimension = new SheetDimension { Reference = "A1" };
                return;
            }

            worksheet.SheetDimension = new SheetDimension
            {
                Reference = ExcelAddress.Cell(minRow, minCol) + ":" + ExcelAddress.Cell(maxRow, maxCol),
            };
        }

        /// <summary>
        /// Drops the cached calculation chain and asks Excel for a full
        /// recalculation on load, so formulas depending on written cells update.
        /// </summary>
        private static void ForceRecalculation(WorkbookPart wbPart)
        {
            if (wbPart.CalculationChainPart != null)
                wbPart.DeletePart(wbPart.CalculationChainPart);

            var calc = wbPart.Workbook.CalculationProperties ?? new CalculationProperties();
            calc.FullCalculationOnLoad = true;
            calc.ForceFullCalculation = true;
            wbPart.Workbook.CalculationProperties = calc;
        }

        /// <summary>
        /// Supplies date number-format styles on demand, and leaves a cell's
        /// existing style alone when it already formats as a date.
        /// </summary>
        private sealed class DateStyleProvider
        {
            // Built-in number format ids that Excel renders as dates/times.
            private static readonly HashSet<uint> BuiltInDateFormats =
                new(new uint[] { 14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47 });

            private readonly WorkbookPart _wbPart;
            private uint? _dateOnlyStyle;
            private uint? _dateTimeStyle;

            public DateStyleProvider(WorkbookPart wbPart) => _wbPart = wbPart;

            public void ApplyDateFormat(Cell cell, DateTime value)
            {
                if (HasDateFormat(cell)) return;

                var wantsTime = value.TimeOfDay != TimeSpan.Zero;
                cell.StyleIndex = wantsTime
                    ? _dateTimeStyle ??= CreateStyle(22)
                    : _dateOnlyStyle ??= CreateStyle(14);
            }

            /// <summary>True when the cell's current style already renders as a date.</summary>
            private bool HasDateFormat(Cell cell)
            {
                var styleIndex = cell.StyleIndex?.Value;
                if (styleIndex is null) return false;

                var stylesheet = _wbPart.WorkbookStylesPart?.Stylesheet;
                var formats = stylesheet?.CellFormats;
                if (formats is null) return false;

                var list = formats.Elements<CellFormat>().ToList();
                if (styleIndex.Value >= list.Count) return false;

                var numFmtId = list[(int)styleIndex.Value].NumberFormatId?.Value;
                if (numFmtId is null) return false;

                if (BuiltInDateFormats.Contains(numFmtId.Value)) return true;

                // Custom format — inspect its code for date/time tokens.
                var custom = stylesheet.NumberingFormats?
                    .Elements<NumberingFormat>()
                    .FirstOrDefault(n => n.NumberFormatId?.Value == numFmtId.Value);

                var code = custom?.FormatCode?.Value;
                if (string.IsNullOrEmpty(code)) return false;

                return code.IndexOfAny(new[] { 'y', 'd', 'h', 's' }) >= 0
                       || code.Contains("mm", StringComparison.Ordinal);
            }

            /// <summary>Appends a cell format using a built-in date number format.</summary>
            private uint CreateStyle(uint numberFormatId)
            {
                var stylesPart = _wbPart.WorkbookStylesPart;
                if (stylesPart is null)
                {
                    stylesPart = _wbPart.AddNewPart<WorkbookStylesPart>();
                    stylesPart.Stylesheet = new Stylesheet();
                }

                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();

                // A valid stylesheet needs at least one font/fill/border/cellStyleXf.
                stylesheet.Fonts ??= new Fonts(new Font()) { Count = 1U };
                stylesheet.Fills ??= new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })) { Count = 1U };
                stylesheet.Borders ??= new Borders(new Border()) { Count = 1U };
                stylesheet.CellStyleFormats ??= new CellStyleFormats(new CellFormat()) { Count = 1U };
                stylesheet.CellFormats ??= new CellFormats(new CellFormat()) { Count = 1U };

                stylesheet.CellFormats.Append(new CellFormat
                {
                    NumberFormatId = numberFormatId,
                    ApplyNumberFormat = true,
                    FormatId = 0U,
                });

                var index = (uint)(stylesheet.CellFormats.Elements<CellFormat>().Count() - 1);
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Elements<CellFormat>().Count();
                stylesheet.Save();

                return index;
            }
        }
    }
}




