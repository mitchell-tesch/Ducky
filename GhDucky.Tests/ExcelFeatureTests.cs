using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using GhDucky.Services;
using GhDucky.Utils;
using Xunit;

namespace GhDucky.Tests
{
    /// <summary>
    /// Covers the second wave of Excel features: macro-enabled workbook
    /// creation, clearing, uncached-formula detection, sheet-name validation
    /// and template rendering.
    /// </summary>
    public sealed class ExcelFeatureTests : IDisposable
    {
        private readonly string _dir;
        private readonly string _path;

        public ExcelFeatureTests()
        {
            _dir = Path.Combine(Path.GetTempPath(), "DuckyExcelFeatureTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
            _path = Path.Combine(_dir, "book.xlsx");
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, true); } catch { /* best effort */ }
        }

        // ---- helpers ---------------------------------------------------------

        private static ExcelWriteResult Write(
            string file, string range, object[][] grid,
            ExcelWriteMode mode = ExcelWriteMode.InPlace, bool formulas = true)
        {
            Assert.True(ExcelAddress.TryParse(range, out var address, out var error), error);
            return ExcelWriter.Write(file, address, grid, mode, formulas);
        }

        private static ExcelReadResult Read(
            string file, string range, bool expand = false, bool formulas = false)
        {
            Assert.True(ExcelAddress.TryParse(range, out var address, out var error), error);
            return ExcelReader.Read(file, address, expand, formulas);
        }

        private static object[][] Grid(params object[][] rows) => rows;

        // ---- 1. macro-enabled workbooks -------------------------------------

        [Fact]
        public void NewXlsmWorkbook_UsesMacroEnabledContentType()
        {
            var xlsm = Path.Combine(_dir, "macro.xlsm");
            Write(xlsm, "Data!A1", Grid(new object[] { "hello" }), ExcelWriteMode.NewWorkbook);

            using var zip = ZipFile.OpenRead(xlsm);
            var entry = zip.GetEntry("[Content_Types].xml");
            Assert.NotNull(entry);

            using var reader = new StreamReader(entry.Open());
            var contentTypes = reader.ReadToEnd();

            // Macro-enabled workbooks use the vnd.ms-excel main part type; the
            // plain spreadsheetml type would make Excel complain that the
            // extension does not match the format.
            Assert.Contains("application/vnd.ms-excel.sheet.macroEnabled.main+xml", contentTypes);
            Assert.DoesNotContain("spreadsheetml.sheet.main+xml", contentTypes);
        }

        [Fact]
        public void NewXlsxWorkbook_UsesPlainWorkbookContentType()
        {
            Write(_path, "Data!A1", Grid(new object[] { "hello" }), ExcelWriteMode.NewWorkbook);

            using var zip = ZipFile.OpenRead(_path);
            using var reader = new StreamReader(zip.GetEntry("[Content_Types].xml")!.Open());
            var contentTypes = reader.ReadToEnd();

            Assert.Contains("spreadsheetml.sheet.main+xml", contentTypes);
            Assert.DoesNotContain("macroEnabled", contentTypes);
        }

        [Fact]
        public void XlsmRoundTripsThroughInPlaceWrite()
        {
            var xlsm = Path.Combine(_dir, "macro.xlsm");
            Write(xlsm, "Data!A1", Grid(new object[] { "one", "two" }), ExcelWriteMode.NewWorkbook);
            Write(xlsm, "Data!B1", Grid(new object[] { "TWO" }));

            var result = Read(xlsm, "Data!A1:B1");
            Assert.Equal("one", result.Grid[0][0]);
            Assert.Equal("TWO", result.Grid[0][1]);
        }

        // ---- 2. sheet name validation ---------------------------------------

        [Theory]
        [InlineData("Has/Slash")]
        [InlineData("Has\\Backslash")]
        [InlineData("Has:Colon")]
        [InlineData("Has?Question")]
        [InlineData("Has*Star")]
        [InlineData("Has[Bracket]")]
        public void InvalidSheetNames_AreRejected(string sheetName)
        {
            var address = new ExcelAddressBuilder(sheetName).Build();
            Assert.Throws<ArgumentException>(() =>
                ExcelWriter.Write(_path, address, Grid(new object[] { "x" }),
                    ExcelWriteMode.NewWorkbook, true));
        }

        [Fact]
        public void OverlongSheetName_IsRejected()
        {
            var address = new ExcelAddressBuilder(new string('a', 32)).Build();
            var ex = Assert.Throws<ArgumentException>(() =>
                ExcelWriter.Write(_path, address, Grid(new object[] { "x" }),
                    ExcelWriteMode.NewWorkbook, true));
            Assert.Contains("31", ex.Message);
        }

        /// <summary>
        /// Builds a sheet-qualified address without going through the A1 parser,
        /// which would itself reject some of these names.
        /// </summary>
        private sealed class ExcelAddressBuilder
        {
            private readonly string _sheet;
            public ExcelAddressBuilder(string sheet) => _sheet = sheet;

            public ExcelAddress Build()
            {
                Assert.True(ExcelAddress.TryParse("A1", out var address, out _));
                return address.WithSheet(_sheet);
            }
        }

        // ---- 3. clearing -----------------------------------------------------

        [Fact]
        public void ClearValues_RemovesContentButKeepsNeighbours()
        {
            Write(_path, "Data!A1", Grid(
                new object[] { "a", "b", "c" },
                new object[] { "d", "e", "f" }),
                ExcelWriteMode.NewWorkbook);

            Assert.True(ExcelAddress.TryParse("Data!B1:B2", out var address, out _));
            var result = ExcelWriter.Clear(_path, address, ExcelClearMode.Values);

            Assert.Equal(2, result.CellsWritten);
            Assert.Equal("Data!B1:B2", result.Address);

            var read = Read(_path, "Data!A1:C2");
            Assert.Equal("a", read.Grid[0][0]);
            Assert.Null(read.Grid[0][1]);
            Assert.Equal("c", read.Grid[0][2]);
            Assert.Equal("d", read.Grid[1][0]);
            Assert.Null(read.Grid[1][1]);
            Assert.Equal("f", read.Grid[1][2]);
        }

        [Fact]
        public void ClearSingleCell_DoesNotAutoExpand()
        {
            Write(_path, "Data!A1", Grid(
                new object[] { "a", "b" },
                new object[] { "c", "d" }),
                ExcelWriteMode.NewWorkbook);

            Assert.True(ExcelAddress.TryParse("Data!A1", out var address, out _));
            var result = ExcelWriter.Clear(_path, address, ExcelClearMode.Values);

            Assert.Equal(1, result.CellsWritten);

            var read = Read(_path, "Data!A1:B2");
            Assert.Null(read.Grid[0][0]);
            Assert.Equal("b", read.Grid[0][1]);
            Assert.Equal("c", read.Grid[1][0]);
            Assert.Equal("d", read.Grid[1][1]);
        }

        [Fact]
        public void ClearEntireSheet_LeavesOtherSheetsAlone()
        {
            Write(_path, "Keep!A1", Grid(new object[] { "survivor" }), ExcelWriteMode.NewWorkbook);
            Write(_path, "Wipe!A1", Grid(
                new object[] { "x", "y" },
                new object[] { "z", "w" }));

            Assert.True(ExcelAddress.TryParse("Wipe!A1", out var address, out _));
            var result = ExcelWriter.Clear(_path, address, ExcelClearMode.EntireSheet);

            Assert.Equal(4, result.CellsWritten);

            Assert.Equal("survivor", Read(_path, "Keep!A1:A1").Grid[0][0]);

            var wiped = Read(_path, "Wipe!A1", expand: true);
            Assert.Equal(0, wiped.RowCount);
        }

        [Fact]
        public void ClearValuesAndFormatting_RemovesTheCells()
        {
            Write(_path, "Data!A1", Grid(new object[] { "a", "b" }), ExcelWriteMode.NewWorkbook);

            Assert.True(ExcelAddress.TryParse("Data!A1:B1", out var address, out _));
            ExcelWriter.Clear(_path, address, ExcelClearMode.ValuesAndFormatting);

            using var zip = ZipFile.OpenRead(_path);
            using var reader = new StreamReader(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open());
            var sheetXml = reader.ReadToEnd();

            // Both cells, and the now-empty row, should be gone from the XML.
            Assert.DoesNotContain("r=\"A1\"", sheetXml);
            Assert.DoesNotContain("r=\"B1\"", sheetXml);
        }

        [Fact]
        public void ClearOnMissingSheet_Throws()
        {
            Write(_path, "Data!A1", Grid(new object[] { "a" }), ExcelWriteMode.NewWorkbook);

            Assert.True(ExcelAddress.TryParse("Nope!A1", out var address, out _));
            Assert.Throws<ArgumentException>(() =>
                ExcelWriter.Clear(_path, address, ExcelClearMode.Values));
        }

        // ---- 4. uncached formula detection -----------------------------------

        [Fact]
        public void UncachedFormula_IsReportedOnRead()
        {
            Write(_path, "Data!A1", Grid(new object[] { 2.0, 3.0 }), ExcelWriteMode.NewWorkbook);
            Write(_path, "Data!C1", Grid(new object[] { "=A1*B1" }));

            var result = Read(_path, "Data!A1:C1");

            // The formula has no cached <v>, so it reads as null...
            Assert.Null(result.Grid[0][2]);

            // ...and the reader tells the component exactly that.
            Assert.Equal(1, result.UncachedFormulaCount);
            Assert.Equal("Data!C1", result.FirstUncachedFormula);
        }

        [Fact]
        public void CachedFormula_IsNotReportedAsUncached()
        {
            // Hand-build a workbook whose formula carries a cached result,
            // mimicking a file Excel has already recalculated.
            var file = Path.Combine(_dir, "cached.xlsx");
            WriteWorkbookWithCachedFormula(file);

            var result = Read(file, "Data!A1:B1");

            Assert.Equal(84.0, result.Grid[0][1]);
            Assert.Equal(0, result.UncachedFormulaCount);
            Assert.Null(result.FirstUncachedFormula);
        }

        [Fact]
        public void NoFormulas_MeansNoWarningState()
        {
            Write(_path, "Data!A1", Grid(new object[] { "a", null, "c" }), ExcelWriteMode.NewWorkbook);

            var result = Read(_path, "Data!A1:C1");

            Assert.Equal(0, result.UncachedFormulaCount);
            Assert.Null(result.FirstUncachedFormula);
        }

        [Fact]
        public void ReadingFormulaText_SuppressesTheUncachedCount()
        {
            Write(_path, "Data!A1", Grid(new object[] { 2.0 }), ExcelWriteMode.NewWorkbook);
            Write(_path, "Data!B1", Grid(new object[] { "=A1*2" }));

            var result = Read(_path, "Data!A1:B1", formulas: true);

            Assert.Equal("=A1*2", result.Grid[0][1]);
            Assert.Equal(0, result.UncachedFormulaCount);
        }

        private static void WriteWorkbookWithCachedFormula(string file)
        {
            using var zip = ZipFile.Open(file, ZipArchiveMode.Create);

            void Add(string name, string xml)
            {
                using var s = new StreamWriter(zip.CreateEntry(name).Open(), new System.Text.UTF8Encoding(false));
                s.Write(xml);
            }

            Add("[Content_Types].xml",
                """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>""");

            Add("_rels/.rels",
                """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""");

            Add("xl/workbook.xml",
                """<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>""");

            Add("xl/_rels/workbook.xml.rels",
                """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""");

            Add("xl/worksheets/sheet1.xml",
                """<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1"><v>42</v></c><c r="B1"><f>A1*2</f><v>84</v></c></row></sheetData></worksheet>""");
        }

        // ---- 5. template rendering ------------------------------------------

        /// <summary>
        /// Builds a template workbook whose text lives in the shared string
        /// table, which is what Excel itself produces. MiniExcel's template
        /// engine reads placeholders from shared strings, so a fixture written
        /// with inline strings would not exercise the real path.
        /// </summary>
        private static void WriteTemplate(string file, string[][] rows)
        {
            var strings = new List<string>();
            var sheetRows = new System.Text.StringBuilder();

            for (var r = 0; r < rows.Length; r++)
            {
                sheetRows.Append($"<row r=\"{r + 1}\">");
                for (var c = 0; c < rows[r].Length; c++)
                {
                    var text = rows[r][c];
                    if (text is null) continue;

                    var index = strings.IndexOf(text);
                    if (index < 0)
                    {
                        strings.Add(text);
                        index = strings.Count - 1;
                    }

                    var reference = ExcelAddress.Cell(r + 1, c + 1);
                    sheetRows.Append($"<c r=\"{reference}\" t=\"s\"><v>{index}</v></c>");
                }
                sheetRows.Append("</row>");
            }

            var sst = new System.Text.StringBuilder();
            sst.Append($"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{strings.Count}\" uniqueCount=\"{strings.Count}\">");
            foreach (var s in strings)
                sst.Append("<si><t xml:space=\"preserve\">").Append(System.Security.SecurityElement.Escape(s)).Append("</t></si>");
            sst.Append("</sst>");

            if (File.Exists(file)) File.Delete(file);
            using var zip = ZipFile.Open(file, ZipArchiveMode.Create);

            void Add(string name, string xml)
            {
                using var s = new StreamWriter(zip.CreateEntry(name).Open(), new System.Text.UTF8Encoding(false));
                s.Write(xml);
            }

            Add("[Content_Types].xml",
                """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>""");

            Add("_rels/.rels",
                """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""");

            Add("xl/workbook.xml",
                """<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>""");

            Add("xl/_rels/workbook.xml.rels",
                """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>""");

            Add("xl/sharedStrings.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + sst);

            Add("xl/worksheets/sheet1.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                + sheetRows + "</sheetData></worksheet>");
        }

        [Fact]
        public void Template_FillsScalarPlaceholders()
        {
            var template = Path.Combine(_dir, "template.xlsx");
            var output = Path.Combine(_dir, "report.xlsx");

            WriteTemplate(template, new[]
            {
                new[] { "{{title}}" },
                new[] { "{{author}}" },
            });

            var values = ExcelTemplate.BuildValues(
                new[] { "title", "author" },
                new IReadOnlyList<object>[]
                {
                    new object[] { "Quarterly Report" },
                    new object[] { "Ducky" },
                });

            ExcelTemplate.Fill(template, output, values, true);

            var result = Read(output, "Sheet1!A1:A2");
            Assert.Equal("Quarterly Report", result.Grid[0][0]);
            Assert.Equal("Ducky", result.Grid[1][0]);
        }

        [Fact]
        public void Template_ExpandsRepeatingCollections()
        {
            var template = Path.Combine(_dir, "template.xlsx");
            var output = Path.Combine(_dir, "report.xlsx");

            WriteTemplate(template, new[]
            {
                new[] { "Name", "Qty" },
                new[] { "{{rows.name}}", "{{rows.qty}}" },
            });

            var values = ExcelTemplate.BuildValues(
                new[] { "rows.name", "rows.qty" },
                new IReadOnlyList<object>[]
                {
                    new object[] { "alpha", "beta", "gamma" },
                    new object[] { 1.0, 2.0, 3.0 },
                });

            ExcelTemplate.Fill(template, output, values, true);

            var result = Read(output, "Sheet1!A1", expand: true);

            // Header row plus one row per collection item.
            Assert.Equal(4, result.RowCount);
            Assert.Equal("Name", result.Grid[0][0]);
            Assert.Equal("alpha", result.Grid[1][0]);
            Assert.Equal(1.0, Convert.ToDouble(result.Grid[1][1]));
            Assert.Equal("beta", result.Grid[2][0]);
            Assert.Equal(2.0, Convert.ToDouble(result.Grid[2][1]));
            Assert.Equal("gamma", result.Grid[3][0]);
            Assert.Equal(3.0, Convert.ToDouble(result.Grid[3][1]));
        }

        [Fact]
        public void BuildValues_GroupsDottedKeysIntoOneCollection()
        {
            var values = ExcelTemplate.BuildValues(
                new[] { "title", "rows.name", "rows.qty" },
                new IReadOnlyList<object>[]
                {
                    new object[] { "T" },
                    new object[] { "a", "b" },
                    new object[] { 1.0, 2.0 },
                });

            Assert.Equal("T", values["title"]);

            var rows = Assert.IsAssignableFrom<List<IDictionary<string, object>>>(values["rows"]);
            Assert.Equal(2, rows.Count);
            Assert.Equal("a", rows[0]["name"]);
            Assert.Equal(1.0, rows[0]["qty"]);
            Assert.Equal("b", rows[1]["name"]);
            Assert.Equal(2.0, rows[1]["qty"]);
        }

        [Fact]
        public void BuildValues_PadsRaggedCollectionColumns()
        {
            var values = ExcelTemplate.BuildValues(
                new[] { "rows.a", "rows.b" },
                new IReadOnlyList<object>[]
                {
                    new object[] { 1.0, 2.0, 3.0 },
                    new object[] { "x" },
                });

            var rows = Assert.IsAssignableFrom<List<IDictionary<string, object>>>(values["rows"]);
            Assert.Equal(3, rows.Count);
            Assert.Equal("x", rows[0]["b"]);
            Assert.Null(rows[1]["b"]);
            Assert.Null(rows[2]["b"]);
        }

        [Fact]
        public void Template_RefusesToOverwriteItself()
        {
            var template = Path.Combine(_dir, "template.xlsx");
            WriteTemplate(template, new[] { new[] { "{{title}}" } });

            var values = ExcelTemplate.BuildValues(new[] { "title" },
                new IReadOnlyList<object>[] { new object[] { "x" } });

            var ex = Assert.Throws<InvalidOperationException>(() =>
                ExcelTemplate.Fill(template, template, values, true));
            Assert.Contains("differ", ex.Message);
        }

        [Fact]
        public void Template_RespectsOverwriteFlag()
        {
            var template = Path.Combine(_dir, "template.xlsx");
            var output = Path.Combine(_dir, "report.xlsx");

            WriteTemplate(template, new[] { new[] { "{{title}}" } });
            Write(output, "Sheet1!A1", Grid(new object[] { "existing" }),
                ExcelWriteMode.NewWorkbook, formulas: false);

            var values = ExcelTemplate.BuildValues(new[] { "title" },
                new IReadOnlyList<object>[] { new object[] { "x" } });

            Assert.Throws<InvalidOperationException>(() =>
                ExcelTemplate.Fill(template, output, values, false));
        }

        [Fact]
        public void Template_MissingTemplateThrows()
        {
            var values = ExcelTemplate.BuildValues(new[] { "title" },
                new IReadOnlyList<object>[] { new object[] { "x" } });

            Assert.Throws<FileNotFoundException>(() =>
                ExcelTemplate.Fill(Path.Combine(_dir, "nope.xlsx"), _path, values, true));
        }
    }
}



