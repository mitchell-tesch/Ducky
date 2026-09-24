using System;
using System.IO;
using System.Linq;
using GhDucky.Services;
using GhDucky.Utils;
using Xunit;

namespace GhDucky.Tests
{
    /// <summary>
    /// End-to-end tests over a real xlsx on disk. These are the tests that
    /// matter most for the Excel components: they verify that an in-place write
    /// genuinely leaves the rest of the workbook alone, which is the behaviour
    /// MiniExcel could not provide and the reason the OpenXML SDK is used on the
    /// write path.
    /// </summary>
    public sealed class ExcelRoundTripTests : IDisposable
    {
        private readonly string _dir;
        private readonly string _path;

        public ExcelRoundTripTests()
        {
            _dir = Path.Combine(Path.GetTempPath(), "DuckyExcelTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
            _path = Path.Combine(_dir, "book.xlsx");
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, true); } catch { /* best effort */ }
        }

        // ---- helpers ---------------------------------------------------------

        private ExcelWriteResult Write(
            string range, object[][] grid,
            ExcelWriteMode mode = ExcelWriteMode.InPlace,
            bool formulas = true)
        {
            Assert.True(ExcelAddress.TryParse(range, out var address, out var error), error);
            return ExcelWriter.Write(_path, address, grid, mode, formulas);
        }

        private ExcelReadResult Read(string range, bool expand = false, bool formulas = false)
        {
            Assert.True(ExcelAddress.TryParse(range, out var address, out var error), error);
            return ExcelReader.Read(_path, address, expand, formulas);
        }

        private static object[][] Grid(params object[][] rows) => rows;

        // ---- tests -----------------------------------------------------------

        [Fact]
        public void NewWorkbook_RoundTripsValueTypes()
        {
            var when = new DateTime(2026, 5, 28, 0, 0, 0, DateTimeKind.Unspecified);

            Write("Data!B2", Grid(
                new object[] { "alpha", 42.0, true, when }),
                ExcelWriteMode.NewWorkbook);

            var result = Read("Data!B2:E2");

            Assert.Equal(1, result.RowCount);
            Assert.Equal(4, result.ColumnCount);

            var row = result.Grid[0];
            Assert.Equal("alpha", Assert.IsType<string>(row[0]));
            Assert.Equal(42.0, Assert.IsType<double>(row[1]));
            Assert.True(Assert.IsType<bool>(row[2]));
            Assert.Equal(when, Assert.IsType<DateTime>(row[3]));
        }

        [Fact]
        public void NewWorkbook_CreatesFileAndSheet()
        {
            Write("Report!A1", Grid(new object[] { "hello" }), ExcelWriteMode.NewWorkbook);

            Assert.True(File.Exists(_path));
            Assert.Contains("Report", ExcelReader.GetSheetNames(_path));
        }

        [Fact]
        public void InPlaceWrite_LeavesNeighbouringCellsUntouched()
        {
            Write("Data!A1", Grid(
                new object[] { "a1", "b1", "c1" },
                new object[] { "a2", "b2", "c2" },
                new object[] { "a3", "b3", "c3" }),
                ExcelWriteMode.NewWorkbook);

            // Overwrite only the centre cell.
            Write("Data!B2", Grid(new object[] { "CHANGED" }));

            var result = Read("Data!A1:C3");

            Assert.Equal("CHANGED", result.Grid[1][1]);

            // Everything else must survive verbatim.
            Assert.Equal("a1", result.Grid[0][0]);
            Assert.Equal("b1", result.Grid[0][1]);
            Assert.Equal("c1", result.Grid[0][2]);
            Assert.Equal("a2", result.Grid[1][0]);
            Assert.Equal("c2", result.Grid[1][2]);
            Assert.Equal("a3", result.Grid[2][0]);
            Assert.Equal("b3", result.Grid[2][1]);
            Assert.Equal("c3", result.Grid[2][2]);
        }

        [Fact]
        public void InPlaceWrite_PreservesOtherSheets()
        {
            Write("First!A1", Grid(new object[] { "keep me" }), ExcelWriteMode.NewWorkbook);
            Write("Second!A1", Grid(new object[] { "second" }));

            // Writing to Second must not disturb First.
            Write("Second!A1", Grid(new object[] { "updated" }));

            var names = ExcelReader.GetSheetNames(_path);
            Assert.Contains("First", names);
            Assert.Contains("Second", names);

            Assert.Equal("keep me", Read("First!A1:A1").Grid[0][0]);
            Assert.Equal("updated", Read("Second!A1:A1").Grid[0][0]);
        }

        [Fact]
        public void InPlaceWrite_PreservesUntouchedFormulas()
        {
            Write("Data!A1", Grid(new object[] { 2.0, 3.0 }), ExcelWriteMode.NewWorkbook);
            Write("Data!C1", Grid(new object[] { "=A1*B1" }));

            // Change a value elsewhere; the formula cell must remain a formula.
            Write("Data!A1", Grid(new object[] { 5.0 }));

            var formulas = Read("Data!A1:C1", formulas: true);
            Assert.Equal("=A1*B1", formulas.Grid[0][2]);
            Assert.Equal(5.0, formulas.Grid[0][0]);
        }

        [Fact]
        public void Formulas_AreWrittenAsFormulasWhenEnabled()
        {
            Write("Data!A1", Grid(new object[] { 4.0, "=A1*10" }), ExcelWriteMode.NewWorkbook);

            var asText = Read("Data!A1:B1", formulas: true);
            Assert.Equal("=A1*10", asText.Grid[0][1]);
        }

        [Fact]
        public void Formulas_AreWrittenAsLiteralTextWhenDisabled()
        {
            Write("Data!A1", Grid(new object[] { "=A1*10" }), ExcelWriteMode.NewWorkbook, formulas: false);

            var result = Read("Data!A1:A1");
            Assert.Equal("=A1*10", result.Grid[0][0]);
        }

        [Fact]
        public void NullValue_ClearsTheCell()
        {
            Write("Data!A1", Grid(new object[] { "x", "y", "z" }), ExcelWriteMode.NewWorkbook);
            Write("Data!B1", Grid(new object[] { null }));

            var result = Read("Data!A1:C1");
            Assert.Equal("x", result.Grid[0][0]);
            Assert.Null(result.Grid[0][1]);
            Assert.Equal("z", result.Grid[0][2]);
        }

        [Fact]
        public void AnchorWrite_AutoExpandsToFitTheData()
        {
            var result = Write("Data!B2", Grid(
                new object[] { 1.0, 2.0, 3.0 },
                new object[] { 4.0, 5.0, 6.0 }),
                ExcelWriteMode.NewWorkbook);

            Assert.Equal("Data!B2:D3", result.Address);
            Assert.Equal(6, result.CellsWritten);
            Assert.False(result.Clipped);
        }

        [Fact]
        public void ExplicitRange_ClipsOversizedData()
        {
            var result = Write("Data!A1:B2", Grid(
                new object[] { 1.0, 2.0, 3.0 },
                new object[] { 4.0, 5.0, 6.0 },
                new object[] { 7.0, 8.0, 9.0 }),
                ExcelWriteMode.NewWorkbook);

            Assert.True(result.Clipped);
            Assert.Equal("Data!A1:B2", result.Address);
            Assert.Equal(4, result.CellsWritten);

            var read = Read("Data!A1:C3");
            Assert.Equal(2, read.RowCount);
            Assert.Equal(2, read.ColumnCount);
        }

        [Fact]
        public void ExpandingRead_StopsAtABlankRow()
        {
            Write("Data!B2", Grid(
                new object[] { "a", "b" },
                new object[] { "c", "d" }),
                ExcelWriteMode.NewWorkbook);

            // Stray data further down, separated by a blank row.
            Write("Data!B6", Grid(new object[] { "stray" }));

            var expanded = Read("Data!B2", expand: true);

            Assert.Equal(2, expanded.RowCount);
            Assert.Equal(2, expanded.ColumnCount);
            Assert.Equal("Data!B2:C3", expanded.Address);
        }

        [Fact]
        public void ExpandingRead_StopsAtABlankColumn()
        {
            Write("Data!A1", Grid(
                new object[] { "a", "b", null, "far" },
                new object[] { "c", "d", null, "away" }),
                ExcelWriteMode.NewWorkbook);

            var expanded = Read("Data!A1", expand: true);

            Assert.Equal(2, expanded.RowCount);
            Assert.Equal(2, expanded.ColumnCount);
            Assert.Equal("Data!A1:B2", expanded.Address);
        }

        [Fact]
        public void NonExpandingAnchorRead_ReturnsASingleCell()
        {
            Write("Data!A1", Grid(
                new object[] { "a", "b" },
                new object[] { "c", "d" }),
                ExcelWriteMode.NewWorkbook);

            var single = Read("Data!A1");

            Assert.Equal(1, single.RowCount);
            Assert.Equal(1, single.ColumnCount);
            Assert.Equal("a", single.Grid[0][0]);
        }

        [Fact]
        public void OverwriteSheetMode_ClearsOnlyTheTargetSheet()
        {
            Write("Keep!A1", Grid(new object[] { "survivor" }), ExcelWriteMode.NewWorkbook);
            Write("Wipe!A1", Grid(
                new object[] { "old1", "old2", "old3" },
                new object[] { "old4", "old5", "old6" }));

            Write("Wipe!A1", Grid(new object[] { "fresh" }), ExcelWriteMode.OverwriteSheet);

            Assert.Equal("survivor", Read("Keep!A1:A1").Grid[0][0]);

            var wiped = Read("Wipe!A1", expand: true);
            Assert.Equal(1, wiped.RowCount);
            Assert.Equal(1, wiped.ColumnCount);
            Assert.Equal("fresh", wiped.Grid[0][0]);
        }

        [Fact]
        public void NewWorkbookMode_ReplacesEverything()
        {
            Write("Old!A1", Grid(new object[] { "gone" }), ExcelWriteMode.NewWorkbook);
            Write("Fresh!A1", Grid(new object[] { "here" }), ExcelWriteMode.NewWorkbook);

            var names = ExcelReader.GetSheetNames(_path);
            Assert.Contains("Fresh", names);
            Assert.DoesNotContain("Old", names);
        }

        [Fact]
        public void SheetNameMatching_IsCaseInsensitive()
        {
            Write("Data!A1", Grid(new object[] { "value" }), ExcelWriteMode.NewWorkbook);

            var result = Read("data!A1:A1");
            Assert.Equal("value", result.Grid[0][0]);
        }

        [Fact]
        public void DefinedNames_AreDiscoverable()
        {
            Write("Data!A1", Grid(new object[] { "x" }), ExcelWriteMode.NewWorkbook);

            // A freshly created workbook has no defined names; the reader should
            // return an empty list rather than throwing.
            var names = ExcelWorkbookMetadata.ReadDefinedNames(_path);
            Assert.NotNull(names);
            Assert.Empty(names);
        }

        [Fact]
        public void GetSheets_ReportsUsedRangeAndVisibility()
        {
            Write("Data!B2", Grid(
                new object[] { 1.0, 2.0 },
                new object[] { 3.0, 4.0 }),
                ExcelWriteMode.NewWorkbook);

            var sheets = ExcelReader.GetSheets(_path);
            var data = sheets.Single(s => s.Name == "Data");

            Assert.False(data.Hidden);
            Assert.Equal("B2:C3", data.UsedRange);
            Assert.Equal(2, data.RowCount);
            Assert.Equal(2, data.ColumnCount);
        }

        [Fact]
        public void GetSheets_HandlesAnEmptySheet()
        {
            Write("Data!B2", Grid(new object[] { 1.0 }), ExcelWriteMode.NewWorkbook);

            Assert.True(ExcelAddress.TryParse("Data!B2", out var address, out var error), error);
            ExcelWriter.Clear(_path, address, ExcelClearMode.EntireSheet);

            // An emptied sheet declares a single-cell dimension (ref="A1"), which
            // used to blow up the inspection path.
            var data = ExcelReader.GetSheets(_path).Single(s => s.Name == "Data");

            Assert.Equal(string.Empty, data.UsedRange);
            Assert.Equal(0, data.RowCount);
            Assert.Equal(0, data.ColumnCount);
        }

        [Fact]
        public void ReadCache_IsInvalidatedByAWrite()
        {
            Write("Data!A1", Grid(new object[] { "first" }), ExcelWriteMode.NewWorkbook);

            Assert.True(ExcelAddress.TryParse("Data!A1:A1", out var address, out _));

            var before = ExcelReadCache.GetOrRead(_path, address, false, false, true);
            Assert.Equal("first", before.Grid[0][0]);

            Write("Data!A1", Grid(new object[] { "second" }));

            var after = ExcelReadCache.GetOrRead(_path, address, false, false, true);
            Assert.Equal("second", after.Grid[0][0]);
        }
    }
}

