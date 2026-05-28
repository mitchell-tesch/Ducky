using System;
using System.Collections.Generic;
using DuckDB.NET.Data;
using GhDucky.Parameters;
using GhDucky.Utils;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;

namespace GhDucky.Components.Import
{
    public class ImportDataTreeComponent : DuckyComponentBase
    {
        public ImportDataTreeComponent()
            : base(
                "Ducky Import Data Tree",
                "DuckyTree",
                "Loads a Grasshopper data tree into a Database table using the high-performance Appender API. " +
                "Each branch represents one column; items in a branch are the column's row values.",
                "Ducky",
                "2 | Import")
        {
        }

        public override Guid ComponentGuid => new Guid("5e6b91ce-5d7c-4e4f-9a8b-86d0a5cae8b8");

        public override GH_Exposure Exposure => GH_Exposure.primary;

        protected override System.Drawing.Bitmap Icon => IconFactory.Build("🌳", IconFactory.Grasshopper);

        private int _inImport;
        private int _inDatabase;
        private int _inData;
        private int _inColumns;
        private int _inTypes;
        private int _inTable;
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
            _inData = pManager.AddGenericParameter("Data", "D",
                "Data tree. Each branch is one column; items are row values.",
                GH_ParamAccess.tree);
            _inColumns = pManager.AddTextParameter("Columns", "C",
                "Column names parallel to the branches of Data. Missing names default to col_0, col_1, ...",
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

            pManager[_inColumns].Optional = true;
            pManager[_inTypes].Optional = true;
            pManager[_inSchema].Optional = true;
        }

        private int _outDatabase;
        private int _outTable;
        private int _outRows;

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            _outDatabase = pManager.AddParameter(new ParamDuckyDbConnection(), "Database", "DB",
                "Database connection (passthrough).",
                GH_ParamAccess.item);
            _outTable = pManager.AddTextParameter("Table", "T",
                "Imported table name.",
                GH_ParamAccess.item);
            _outRows = pManager.AddIntegerParameter("Rows", "R",
                "Row count of the table after the import.",
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

            if (!da.GetDataTree(_inData, out GH_Structure<IGH_Goo> tree) || tree == null || tree.IsEmpty)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Data tree is empty.");
                return;
            }

            // Materialize branches as parallel arrays of CLR values.
            var branches = tree.Branches;
            var columnCount = branches.Count;
            if (columnCount == 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Data tree has no branches.");
                return;
            }

            var columnNames = new List<string>();
            var columnTypes = new List<string>();
            var schema = "main";
            var overwrite = true;
            da.GetDataList(_inColumns, columnNames);
            da.GetDataList(_inTypes, columnTypes);
            da.GetData(_inSchema, ref schema);
            da.GetData(_inOverwrite, ref overwrite);

            var columns = new object[columnCount][];
            var rowCount = 0;
            for (var c = 0; c < columnCount; c++)
            {
                var branch = branches[c];
                var col = new object[branch.Count];
                for (var r = 0; r < branch.Count; r++)
                    col[r] = TypeMapping.Unwrap(branch[r]);
                columns[c] = col;
                if (col.Length > rowCount) rowCount = col.Length;
            }

            var names = TypeMapping.ResolveColumnNames(columnNames, columnCount);
            var types = TypeMapping.ResolveColumnTypes(columnTypes, columns);

            var quotedTable = SqlIdentifier.QuoteTable(schema, table);
            var resolvedSchema = string.IsNullOrWhiteSpace(schema) ? "main" : schema;

            try
            {
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
                        // Appending to an existing table: pull its schema so coercion lines up.
                        var existing = ReadTableSchema(conn, resolvedSchema, table);
                        if (existing.Count != columnCount)
                        {
                            AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                                $"Existing table has {existing.Count} columns; data tree has {columnCount} branches.");
                            return;
                        }
                        for (var i = 0; i < columnCount; i++)
                            types[i] = existing[i].Type;
                    }

                    // Pre-build a typed append delegate per column to avoid
                    // the per-cell switch overhead in AppendCell.
                    var appenders = new Action<IDuckDBAppenderRow, object>[columnCount];
                    for (var c = 0; c < columnCount; c++)
                        appenders[c] = BuildCellAppender(types[c]);

                    using (var appender = SqlIdentifier.IsExplicitSchema(schema)
                        ? conn.CreateAppender(schema, table)
                        : conn.CreateAppender(table))
                    {
                        for (var r = 0; r < rowCount; r++)
                        {
                            var row = appender.CreateRow();
                            for (var c = 0; c < columnCount; c++)
                            {
                                var col = columns[c];
                                var raw = r < col.Length ? col[r] : null;
                                appenders[c](row, raw);
                            }
                            row.EndRow();
                        }
                    }

                    // We already know the row count from the data tree —
                    // no need for an extra SELECT COUNT(*) round-trip.
                    da.SetData(_outDatabase, dbConnection);
                    da.SetData(_outTable, table);
                    da.SetData(_outRows, rowCount);
                });
            }
            catch (Exception ex)
            {
                ReportError("Data Tree import failed", ex);
            }
        }


        private static string BuildCreateTable(string quotedTable, List<string> names, List<DuckyColumnType> types)
        {
            var sb = new System.Text.StringBuilder();
            sb.Append("CREATE TABLE ").Append(quotedTable).Append(" (");
            for (var i = 0; i < names.Count; i++)
            {
                if (i > 0) sb.Append(", ");
                sb.Append(SqlIdentifier.Quote(names[i])).Append(' ').Append(types[i].ToSqlText());
            }
            sb.Append(')');
            return sb.ToString();
        }


        private static List<(string Name, DuckyColumnType Type)> ReadTableSchema(DuckDBConnection conn, string schema, string table)
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


        /// <summary>
        /// Returns a pre-built delegate that appends a single cell in the
        /// format the appender expects for the given SQL type in order to avoid
        /// the per-cell switch overhead that <c>AppendCell</c> had.
        /// </summary>
        private static Action<IDuckDBAppenderRow, object> BuildCellAppender(DuckyColumnType sqlType)
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
    }
}
