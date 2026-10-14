using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using DuckDB.NET.Data;
using GhDucky.Parameters;
using GhDucky.Services;
using GhDucky.Utils;
using Grasshopper.Kernel;
namespace GhDucky.Components.Import
{
    /// <summary>
    /// Loads an arbitrary Excel cell range into a DuckDB table.
    /// <para>
    /// Complements <c>Ducky Import Excel</c>, which goes through DuckDB
    /// <c>st_read</c> and can only take a whole sheet. This component uses
    /// the Ducky range reader, so any sub-range, defined name or Excel Table
    /// can be queried in SQL.
    /// </para>
    /// </summary>
    public class ImportExcelRangeComponent : DuckyComponentBase
    {
        public ImportExcelRangeComponent()
            : base(
                "Ducky Import Excel Range",
                "DuckyXlRng",
                "Loads a cell range from an Excel workbook (.xlsx/.xlsm) into a Database table, " +
                "so it can be queried in SQL. Unlike Ducky Import Excel this is not limited to whole " +
                "sheets - any range, defined name or Excel Table can be imported, and no Excel or " +
                "COM install is required.",
                "Ducky",
                "2 | Import")
        {
        }
        public override Guid ComponentGuid => new Guid("c31f7b68-4e0d-42a9-9c85-7a6f2d4b31e0");
        public override GH_Exposure Exposure => GH_Exposure.secondary;
        protected override System.Drawing.Bitmap Icon => IconFactory.Build("📑", IconFactory.ImportFile);
        private int _inImport;
        private int _inDatabase;
        private int _inTable;
        private int _inPath;
        private int _inSheet;
        private int _inRange;
        private int _inHeaders;
        private int _inExpand;
        private int _inColumns;
        private int _inTypes;
        private int _inSchema;
        private int _inOverwrite;
        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            _inImport = pManager.AddBooleanParameter("Import?", "I?",
                "Set to true to perform the import.",
                GH_ParamAccess.item, false);
            _inDatabase = pManager.AddParameter(new ParamDuckyDbConnection(), "Database", "DB",
                "Database connection.",
                GH_ParamAccess.item);
            _inTable = pManager.AddTextParameter("Table", "T",
                "Target table name.",
                GH_ParamAccess.item);
            _inPath = pManager.AddTextParameter("Path", "P",
                "Absolute path to the workbook (.xlsx or .xlsm).",
                GH_ParamAccess.item);
            _inSheet = pManager.AddTextParameter("Sheet", "S",
                "Worksheet name. Leave empty to use the first sheet.",
                GH_ParamAccess.item, string.Empty);
            _inRange = pManager.AddTextParameter("Range", "Rg",
                "Cell or range in A1 notation, optionally sheet-qualified. " +
                "A defined name or Excel Table name may also be given.",
                GH_ParamAccess.item, "A1");
            _inHeaders = pManager.AddBooleanParameter("Headers?", "H?",
                "Treat the first row of the range as column names.",
                GH_ParamAccess.item, true);
            _inExpand = pManager.AddBooleanParameter("Expand?", "E?",
                "When a single anchor cell is given, expand to the enclosing contiguous block.",
                GH_ParamAccess.item, true);
            _inColumns = pManager.AddTextParameter("Columns", "C",
                "Explicit column names, overriding Headers?. Missing names default to col_0, col_1, ...",
                GH_ParamAccess.list);
            _inTypes = pManager.AddTextParameter("Types", "Ty",
                "Optional explicit column types (BOOLEAN, INTEGER, BIGINT, DOUBLE, VARCHAR, TIMESTAMP). " +
                "Defaults to auto-inferred per column.",
                GH_ParamAccess.list);
            _inSchema = pManager.AddTextParameter("Schema", "Sc",
                "Target schema (default: main). Created automatically if missing.",
                GH_ParamAccess.item, "main");
            _inOverwrite = pManager.AddBooleanParameter("Overwrite?", "O?",
                "If true (default), the table is dropped and recreated. otherwise rows are appended.",
                GH_ParamAccess.item, true);
            pManager[_inSheet].Optional = true;
            pManager[_inRange].Optional = true;
            pManager[_inHeaders].Optional = true;
            pManager[_inExpand].Optional = true;
            pManager[_inColumns].Optional = true;
            pManager[_inTypes].Optional = true;
            pManager[_inSchema].Optional = true;
            pManager[_inOverwrite].Optional = true;
        }
        private int _outDatabase;
        private int _outTable;
        private int _outRows;
        private int _outAddress;
        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            _outDatabase = pManager.AddParameter(new ParamDuckyDbConnection(), "Database", "DB",
                "Database connection (passthrough).",
                GH_ParamAccess.item);
            _outTable = pManager.AddTextParameter("Table", "T",
                "Imported table name.",
                GH_ParamAccess.item);
            _outRows = pManager.AddIntegerParameter("Rows", "R",
                "Number of rows imported.",
                GH_ParamAccess.item);
            _outAddress = pManager.AddTextParameter("Address", "A",
                "The Excel address actually read, e.g. Data!B2:D11.",
                GH_ParamAccess.item);
        }
        protected override void SolveInstance(IGH_DataAccess da)
        {
            var import = false;
            da.GetData(_inImport, ref import);
            if (!import)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Import is false; no action taken.");
                return;
            }
            if (!TryGetSession(da, _inDatabase, out var session, out var dbConnection))
                return;
            var table = string.Empty;
            if (!da.GetData(_inTable, ref table) || string.IsNullOrWhiteSpace(table))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Table name required.");
                return;
            }
            var path = string.Empty;
            if (!da.GetData(_inPath, ref path) || string.IsNullOrWhiteSpace(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Path to workbook required.");
                return;
            }
            if (!File.Exists(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Workbook not found: " + path);
                return;
            }
            if (!ExcelFileGuard.IsSupportedExtension(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    "Only .xlsx and .xlsm workbooks are supported. Legacy .xls is not.");
                return;
            }
            if (ExcelFileGuard.IsLockedForReading(path, out var lockReason))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, lockReason);
                return;
            }
            var sheet = string.Empty;
            var range = "A1";
            var headers = true;
            var expand = true;
            var suppliedNames = new List<string>();
            var suppliedTypes = new List<string>();
            var schema = "main";
            var overwrite = true;
            da.GetData(_inSheet, ref sheet);
            da.GetData(_inRange, ref range);
            da.GetData(_inHeaders, ref headers);
            da.GetData(_inExpand, ref expand);
            da.GetDataList(_inColumns, suppliedNames);
            da.GetDataList(_inTypes, suppliedTypes);
            da.GetData(_inSchema, ref schema);
            da.GetData(_inOverwrite, ref overwrite);
            if (string.IsNullOrWhiteSpace(range)) range = "A1";
            try
            {
                if (!TryResolveAddress(path, range, sheet, out var address, out var addressError))
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, addressError);
                    return;
                }
                var result = ExcelReadCache.GetOrRead(path, address, expand, false, true);
                var grid = result.Grid;
                if (grid.Length == 0 || result.ColumnCount == 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "The requested range is empty.");
                    return;
                }
                var columnCount = grid[0].Length;
                // Resolve column names: explicit input wins, then the header row.
                var headerNames = new List<string>(columnCount);
                var dataStart = 0;
                if (headers)
                {
                    for (var c = 0; c < columnCount; c++)
                    {
                        var raw = grid[0][c];
                        var text = raw is null
                            ? null
                            : raw as string ?? Convert.ToString(raw, CultureInfo.InvariantCulture);
                        headerNames.Add(text);
                    }
                    dataStart = 1;
                }
                else
                {
                    for (var c = 0; c < columnCount; c++) headerNames.Add(null);
                }
                for (var c = 0; c < columnCount; c++)
                {
                    if (c < suppliedNames.Count && !string.IsNullOrWhiteSpace(suppliedNames[c]))
                        headerNames[c] = suppliedNames[c];
                }
                var names = DuckDbAppenderFactory.MakeUniqueColumnNames(headerNames);
                var rowCount = grid.Length - dataStart;
                if (rowCount <= 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                        "The range contains only a header row; there are no data rows to import.");
                    return;
                }
                // Pivot the row-major grid into columns for type inference and
                // for the column-oriented appender loop.
                var columns = new object[columnCount][];
                for (var c = 0; c < columnCount; c++)
                {
                    var column = new object[rowCount];
                    for (var r = 0; r < rowCount; r++)
                        column[r] = grid[dataStart + r][c];
                    columns[c] = column;
                }
                var types = TypeMapping.ResolveColumnTypes(suppliedTypes, columns);
                var quotedTable = SqlIdentifier.QuoteTable(schema, table);
                var resolvedSchema = string.IsNullOrWhiteSpace(schema) ? "main" : schema;
                session.Execute(conn =>
                {
                    EnsureSchemaAndDrop(conn, schema, quotedTable, overwrite);
                    if (overwrite || !TableExists(conn, resolvedSchema, table))
                    {
                        using var createCmd = conn.CreateCommand();
                        createCmd.CommandText = BuildCreateTable(quotedTable, names, types);
                        createCmd.ExecuteNonQuery();
                    }
                    else
                    {
                        var existing = ReadTableSchema(conn, resolvedSchema, table);
                        if (existing.Count != columnCount)
                        {
                            AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                                $"Existing table has {existing.Count} columns; the range has {columnCount}.");
                            return;
                        }
                        for (var i = 0; i < columnCount; i++)
                            types[i] = existing[i].Type;
                    }
                    var appenders = new Action<IDuckDBAppenderRow, object>[columnCount];
                    for (var c = 0; c < columnCount; c++)
                        appenders[c] = DuckDbAppenderFactory.BuildCellAppender(types[c]);
                    using (var appender = SqlIdentifier.IsExplicitSchema(schema)
                               ? conn.CreateAppender(schema, table)
                               : conn.CreateAppender(table))
                    {
                        for (var r = 0; r < rowCount; r++)
                        {
                            var row = appender.CreateRow();
                            for (var c = 0; c < columnCount; c++)
                                appenders[c](row, columns[c][r]);
                            row.EndRow();
                        }
                    }
                    da.SetData(_outDatabase, dbConnection);
                    da.SetData(_outTable, table);
                    da.SetData(_outRows, rowCount);
                    da.SetData(_outAddress, result.Address);
                });
                if (result.UncachedFormulaCount > 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                        $"{result.UncachedFormulaCount} formula cell(s) in this range have no cached result " +
                        $"(first: {result.FirstUncachedFormula}) and were imported as NULL. " +
                        "Open the workbook in Excel once so it recalculates, then re-import.");
                }
            }
            catch (Exception ex)
            {
                ReportError("Excel range import failed", ex);
            }
        }
        /// <summary>
        /// Parses the range and resolves defined names / Table names against the
        /// workbook.
        /// </summary>
        private static bool TryResolveAddress(
            string path, string range, string sheet, out ExcelAddress address, out string error)
        {
            if (!ExcelAddress.TryParse(range, out address, out error))
                return false;
            if (address.IsName)
            {
                var reference = ExcelWorkbookMetadata.ResolveName(path, address.Name,
                    string.IsNullOrWhiteSpace(sheet) ? null : sheet);
                if (reference is null)
                {
                    error = $"'{address.Name}' is not a defined name or Table in this workbook.";
                    return false;
                }
                if (!ExcelAddress.TryParse(reference, out address, out error))
                {
                    error = $"Defined name '{range}' resolves to '{reference}', which could not be parsed.";
                    return false;
                }
            }
            if (address.Sheet is null && !string.IsNullOrWhiteSpace(sheet))
                address = address.WithSheet(sheet);
            return true;
        }
        private static string BuildCreateTable(string quotedTable, List<string> names, List<DuckyColumnType> types)
        {
            var sb = new StringBuilder();
            sb.Append("CREATE TABLE ").Append(quotedTable).Append(" (");
            for (var i = 0; i < names.Count; i++)
            {
                if (i > 0) sb.Append(", ");
                sb.Append(SqlIdentifier.Quote(names[i])).Append(' ').Append(types[i].ToSqlText());
            }
            sb.Append(')');
            return sb.ToString();
        }
        private static List<(string Name, DuckyColumnType Type)> ReadTableSchema(
            DuckDBConnection conn, string schema, string table)
        {
            var list = new List<(string Name, DuckyColumnType Type)>();
            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT column_name, data_type FROM information_schema.columns " +
                $"WHERE table_schema = {SqlIdentifier.QuoteLiteral(schema)} " +
                $"AND table_name = {SqlIdentifier.QuoteLiteral(table)} " +
                "ORDER BY ordinal_position";
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                list.Add((reader.GetString(0), DuckyColumnTypeExtensions.ParseOrVarchar(reader.GetString(1))));
            return list;
        }
    }
}