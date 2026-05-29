using System;
using System.Collections.Generic;
using System.Globalization;
using GhDucky.Parameters;
using GhDucky.Services;
using GhDucky.Utils;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;

namespace GhDucky.Components.Import
{
    public class ImportGeometryComponent : DuckyComponentBase
    {
        public ImportGeometryComponent()
            : base(
                "Ducky Import Geometry",
                "DuckyGeo",
                "Writes Rhino geometry to a DuckDB table as a GEOMETRY column (WKB-encoded). " +
                "Auto-loads the spatial extension. Points, curves (tessellated), meshes and breps (auto-meshed) are supported.",
                "Ducky",
                "2 | Import")
        {
        }

        public override Guid ComponentGuid => new Guid("4f9c5bf8-1c2b-4a7e-9f63-3a8d1c6d2a1b");

        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        protected override System.Drawing.Bitmap Icon => IconFactory.Build("🦏", IconFactory.Spatial);

        private int _inImport;
        private int _inDatabase;
        private int _inTable;
        private int _inGeometry;
        private int _inIds;
        private int _inData;
        private int _inColumns;
        private int _inTypes;
        private int _inGeomColumn;
        private int _inSchema;
        private int _inTolerance;
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
            _inGeometry = pManager.AddGeometryParameter("Geometry", "G",
                "Geometry to write.",
                GH_ParamAccess.list);
            _inIds = pManager.AddIntegerParameter("Ids", "Id",
                "Optional integer ids parallel to Geometry. Written to an 'id' column.",
                GH_ParamAccess.list);
            _inData = pManager.AddGenericParameter("Data", "D",
                "Optional data tree of extra columns. Each branch is one column; items are row values parallel to Geometry.",
                GH_ParamAccess.tree);
            _inColumns = pManager.AddTextParameter("Columns", "C",
                "Column names parallel to the branches of Data. Missing names default to col_0, col_1, ...",
                GH_ParamAccess.list);
            _inTypes = pManager.AddTextParameter("Types", "Ty",
                "Optional explicit column types for extra data (BOOLEAN, INTEGER, BIGINT, DOUBLE, VARCHAR, TIMESTAMP). " +
                "Defaults to auto-inferred per column.",
                GH_ParamAccess.list);
            _inGeomColumn = pManager.AddTextParameter("GeomColumn", "Gc",
                "Name for the geometry column.",
                GH_ParamAccess.item, "geom");
            _inTolerance = pManager.AddNumberParameter("Tolerance", "Tol",
                "Tessellation tolerance for non-polyline curves and breps. 0 = use default sampling.",
                GH_ParamAccess.item, 0.01);
            _inSchema = pManager.AddTextParameter("Schema", "Sc",
                "Target schema (default: main). Created automatically if missing.",
                GH_ParamAccess.item, "main");
            _inOverwrite = pManager.AddBooleanParameter("Overwrite?", "O?",
                "If true, the table is dropped and recreated; otherwise rows are appended.",
                GH_ParamAccess.item, true);

            pManager[_inIds].Optional = true;
            pManager[_inData].Optional = true;
            pManager[_inColumns].Optional = true;
            pManager[_inTypes].Optional = true;
            pManager[_inSchema].Optional = true;
            pManager[_inGeomColumn].Optional = true;
            pManager[_inTolerance].Optional = true;
        }

        private int _outDatabase;
        private int _outTable;
        private int _outRows;

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            _outDatabase = pManager.AddParameter(new ParamDuckyDbConnection(), "Database", "DB",
                "Database Connection (passthrough).",
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

            var geometries = new List<IGH_GeometricGoo>();
            if (!da.GetDataList(_inGeometry, geometries) || geometries.Count == 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "No geometry supplied.");
                return;
            }

            var ids = new List<int>();
            da.GetDataList(_inIds, ids);

            // --- Extra data columns (optional) ---
            da.GetDataTree(_inData, out GH_Structure<IGH_Goo> dataTree);
            var hasExtraData = dataTree != null && !dataTree.IsEmpty;

            var extraColumns = (object[][])null;
            var extraNames = new List<string>();
            var extraTypes = new List<DuckyColumnType>();

            if (hasExtraData)
            {
                var branches = dataTree.Branches;
                var columnCount = branches.Count;

                extraColumns = new object[columnCount][];
                for (var c = 0; c < columnCount; c++)
                {
                    var branch = branches[c];
                    var col = new object[branch.Count];
                    for (var r = 0; r < branch.Count; r++)
                        col[r] = TypeMapping.Unwrap(branch[r]);
                    extraColumns[c] = col;
                }

                var suppliedNames = new List<string>();
                var suppliedTypes = new List<string>();
                da.GetDataList(_inColumns, suppliedNames);
                da.GetDataList(_inTypes, suppliedTypes);

                extraNames = TypeMapping.ResolveColumnNames(suppliedNames, columnCount);
                extraTypes = TypeMapping.ResolveColumnTypes(suppliedTypes, extraColumns);
            }

            var geomColumn = "geom";
            da.GetData(_inGeomColumn, ref geomColumn);
            if (string.IsNullOrWhiteSpace(geomColumn)) geomColumn = "geom";

            var tolerance = 0.01;
            da.GetData(_inTolerance, ref tolerance);

            var schema = "main";
            da.GetData(_inSchema, ref schema);

            var overwrite = true;
            da.GetData(_inOverwrite, ref overwrite);

            try
            {
                SpatialExtension.Ensure(session);
            }
            catch (Exception ex)
            {
                ReportError("Could not load spatial extension", ex);
                return;
            }

            // Encode all geometries up-front so a failure aborts before we touch the table.
            var encoded = new byte[geometries.Count][];
            for (var i = 0; i < geometries.Count; i++)
            {
                var raw = UnwrapGeometry(geometries[i]);
                if (raw == null)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                        $"Item {i} is not geometry that can be encoded.");
                    return;
                }
                try
                {
                    encoded[i] = WkbCodec.Encode(raw, tolerance);
                }
                catch (Exception ex)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                        $"Item {i} ({raw.GetType().Name}) could not be encoded: {ex.Message}");
                    return;
                }
            }

            var quotedTable = SqlIdentifier.QuoteTable(schema, table);
            var quotedGeom = SqlIdentifier.Quote(geomColumn);

            try
            {
                session.Execute(conn =>
                {
                    EnsureSchemaAndDrop(conn, schema, quotedTable, overwrite);

                    using (var cmd = conn.CreateCommand())
                    {
                        // For both overwrite and first-time append, ensure the
                        // table exists with the right shape. IF NOT EXISTS makes
                        // subsequent appends a no-op when the table is already
                        // present from an earlier run.
                        cmd.CommandText = BuildCreateTable(quotedTable, quotedGeom, extraNames, extraTypes,
                            createIfNotExists: !overwrite);
                        cmd.ExecuteNonQuery();
                    }

                    // Insert geometry rows inside a transaction for atomicity.
                    using (var tx = conn.BeginTransaction())
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tx;

                        // Build the column names header once (it is the same for every batch).
                        var colNamesHeader = new System.Text.StringBuilder();
                        colNamesHeader.Append("id, ").Append(quotedGeom);
                        if (hasExtraData)
                        {
                            for (var c = 0; c < extraColumns.Length; c++)
                                colNamesHeader.Append(", ").Append(SqlIdentifier.Quote(extraNames[c]));
                        }
                        var colNamesStr = colNamesHeader.ToString();

                        // Build multi-row INSERTs using a VALUES clause with
                        // ST_GeomFromWKB(unhex(...)) so the spatial conversion happens
                        // in bulk inside DuckDB rather than one round-trip per row.
                        const int batchSize = 2000;
                        for (var batchStart = 0; batchStart < encoded.Length; batchStart += batchSize)
                        {
                            var batchEnd = Math.Min(batchStart + batchSize, encoded.Length);

                            var sql = new System.Text.StringBuilder();
                            sql.Append("INSERT INTO ").Append(quotedTable)
                               .Append(" (").Append(colNamesStr).Append(") VALUES ");

                            for (var i = batchStart; i < batchEnd; i++)
                            {
                                if (i > batchStart) sql.Append(", ");

                                var idVal = i < ids.Count ? ids[i] : i;
                                var hexLiteral = SqlIdentifier.QuoteLiteral(Convert.ToHexString(encoded[i]));

                                sql.Append('(')
                                   .Append(idVal.ToString(CultureInfo.InvariantCulture))
                                   .Append(", ST_GeomFromWKB(unhex(")
                                   .Append(hexLiteral)
                                   .Append("))");

                                if (hasExtraData)
                                {
                                    for (var c = 0; c < extraColumns.Length; c++)
                                    {
                                        var raw = i < extraColumns[c].Length ? extraColumns[c][i] : null;
                                        sql.Append(", ").Append(FormatSqlLiteral(raw, extraTypes[c]));
                                    }
                                }

                                sql.Append(')');
                            }

                            cmd.CommandText = sql.ToString();
                            cmd.ExecuteNonQuery();
                        }

                        tx.Commit();
                    }

                    // In overwrite mode we know the exact row count without querying.
                    // In append mode we need to count the full table.
                    var rowCount = overwrite
                        ? encoded.Length
                        : CountRows(conn, quotedTable);

                    da.SetData(_outDatabase, dbConnection);
                    da.SetData(_outTable, table);
                    da.SetData(_outRows, rowCount);
                });
            }
            catch (Exception ex)
            {
                ReportError("Geometry import failed", ex);
            }
        }


        private static object UnwrapGeometry(IGH_GeometricGoo goo)
        {
            if (goo == null) return null;
            return goo.ScriptVariable() ?? goo;
        }


        private static string BuildCreateTable(string quotedTable, string quotedGeom,
            List<string> extraNames, List<DuckyColumnType> extraTypes, bool createIfNotExists = false)
        {
            var sb = new System.Text.StringBuilder();
            sb.Append(createIfNotExists ? "CREATE TABLE IF NOT EXISTS " : "CREATE TABLE ")
              .Append(quotedTable)
              .Append(" (id INTEGER, ").Append(quotedGeom).Append(" GEOMETRY");

            for (var i = 0; i < extraNames.Count; i++)
            {
                sb.Append(", ")
                  .Append(SqlIdentifier.Quote(extraNames[i]))
                  .Append(' ')
                  .Append(extraTypes[i].ToSqlText());
            }

            sb.Append(')');
            return sb.ToString();
        }



        /// <summary>
        /// Formats a CLR value as a SQL literal string suitable for inlining in an INSERT statement.
        /// </summary>
        private static string FormatSqlLiteral(object value, DuckyColumnType sqlType)
        {
            if (value is null) return "NULL";

            var coerced = TypeMapping.CoerceForAppender(value, sqlType);
            if (coerced is null) return "NULL";

            switch (sqlType)
            {
                case DuckyColumnType.Boolean:
                    return (bool)coerced ? "TRUE" : "FALSE";
                case DuckyColumnType.Integer:
                    return ((int)coerced).ToString(CultureInfo.InvariantCulture);
                case DuckyColumnType.BigInt:
                    return ((long)coerced).ToString(CultureInfo.InvariantCulture);
                case DuckyColumnType.Double:
                    return ((double)coerced).ToString("R", CultureInfo.InvariantCulture);
                case DuckyColumnType.Timestamp:
                    var dt = (DateTime)coerced;
                    return SqlIdentifier.QuoteLiteral(
                        dt.ToString("yyyy-MM-dd HH:mm:ss.ffffff", CultureInfo.InvariantCulture));
                case DuckyColumnType.Varchar:
                default:
                    return SqlIdentifier.QuoteLiteral(Convert.ToString(coerced, CultureInfo.InvariantCulture));
            }
        }
    }
}
