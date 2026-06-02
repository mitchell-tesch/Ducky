using System;
using System.Collections.Generic;
using System.IO;
using GhDucky.Parameters;
using GhDucky.Services;
using GhDucky.Utils;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace GhDucky.Components.Query
{
    public class QueryGeometryComponent : DuckyComponentBase
    {
        public QueryGeometryComponent()
            : base(
                "Ducky Query Geometry",
                "DuckyGeoQ",
                "Runs a SQL query that returns a GEOMETRY column and reconstructs Rhino geometry from the WKB payload. " +
                "Other columns are returned as a parallel data tree.",
                "Ducky",
                "3 | Query")
        {
        }

        public override Guid ComponentGuid => new Guid("8d9f4a82-4b1d-4fb1-93f2-2c7d2c5f0a1c");

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        protected override System.Drawing.Bitmap Icon => IconFactory.Build("❓", IconFactory.Spatial);

        private int _inRun;
        private int _inDatabase;
        private int _inQuery;
        private int _inGeomColumn;
        private int _inLimit;
        private int _inTyped;

        private const int DefaultRowLimit = 1_000_000;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            _inRun = pManager.AddBooleanParameter("Run?", "R?",
                "Set to true to execute the query.",
                GH_ParamAccess.item, false);
            _inDatabase = pManager.AddParameter(new ParamDuckyDbConnection(), "Database", "DB",
                "Database connection.",
                GH_ParamAccess.item);
            _inQuery = pManager.AddTextParameter("Query", "SQL",
                "SQL query to execute. SELECT statements return a result tree; The geometry column named by GeomColumn is rewritten to its WKB form automatically.",
                GH_ParamAccess.item);
            _inGeomColumn = pManager.AddTextParameter("GeomColumn", "Gc",
                "Name of the geometry column in the result set.",
                GH_ParamAccess.item, "geom");
            _inLimit = pManager.AddIntegerParameter("Limit", "L",
                "Maximum number of rows to fetch. Defaults to 1,000,000 to guard against excessive memory use; " +
                "set to 0 to return all rows (use with care on large result sets).",
                GH_ParamAccess.item, DefaultRowLimit);
            _inTyped = pManager.AddBooleanParameter("Typed?", "Ty?",
                "If true (default), temporal columns (DATE/TIME/TIMESTAMP) surface as GH_Time so downstream date arithmetic works. " +
                "If false, those values are emitted as ISO 8601 strings.",
                GH_ParamAccess.item, true);

            pManager[_inLimit].Optional = true;
            pManager[_inTyped].Optional = true;
        }

        private int _outGeometry;
        private int _outColumns;
        private int _outData;
        private int _outRows;

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            _outGeometry = pManager.AddGeometryParameter("Geometry", "G",
                "Reconstructed Rhino geometry, one entry per row.",
                GH_ParamAccess.list);
            _outColumns = pManager.AddTextParameter("Columns", "N",
                "Non-geometry column names, parallel to the branches of Data.",
                GH_ParamAccess.list);
            _outData = pManager.AddGenericParameter("Data", "D",
                "Non-geometry result data tree: one branch per column, items are row values.",
                GH_ParamAccess.tree);
            _outRows = pManager.AddIntegerParameter("Rows", "R",
                "Row count of the result set (0 for non-SELECT queries).",
                GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            var run = false;
            da.GetData(_inRun, ref run);
            if (!run)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Run is false; query not executed.");
                return;
            }

            if (!TryGetSession(da, _inDatabase, out var session))
                return;

            string query = null;
            if (!da.GetData(_inQuery, ref query) || string.IsNullOrWhiteSpace(query))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "SQL required.");
                return;
            }

            var geomColumn = "geom";
            da.GetData(_inGeomColumn, ref geomColumn);
            if (string.IsNullOrWhiteSpace(geomColumn))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "GeomColumn required.");
                return;
            }

            var limit = DefaultRowLimit;
            da.GetData(_inLimit, ref limit);
            if (limit < 0) limit = 0;

            var typed = true;
            da.GetData(_inTyped, ref typed);

            try
            {
                SpatialExtension.Ensure(session);
            }
            catch (Exception ex)
            {
                ReportError("Could not load spatial extension", ex);
                return;
            }

            // Wrap the user's query so that the geometry column comes back as a
            // BLOB of WKB bytes, and other columns are passed through unchanged.
            var quotedGeom = SqlIdentifier.Quote(geomColumn);
            var limitClause = limit > 0
                ? $" LIMIT {limit.ToString(System.Globalization.CultureInfo.InvariantCulture)}"
                : string.Empty;
            var wrapped =
                $"SELECT * EXCLUDE ({quotedGeom}), ST_AsWKB({quotedGeom}) AS {quotedGeom} " +
                $"FROM ({query.TrimEnd(';')}) AS __ghducky_sub{limitClause}";

            try
            {
                session.Execute(conn =>
                {
                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = wrapped;
                    using var reader = cmd.ExecuteReader();

                    var columnCount = reader.FieldCount;
                    var geomIndex = -1;
                    var nonGeomColumnNames = new List<string>(columnCount - 1);
                    var nonGeomIndices = new List<int>(columnCount - 1);
                    var nonGeomReaders = new List<Func<System.Data.Common.DbDataReader, int, IGH_Goo>>(columnCount - 1);

                    // DuckDB folds unquoted identifiers to lower-case, so compare
                    // case-insensitively to avoid spurious "no such column" errors
                    // when the user passes e.g. "Geom" while the SELECT returns "geom".
                    for (var i = 0; i < columnCount; i++)
                    {
                        var name = reader.GetName(i);
                        if (geomIndex < 0 &&
                            string.Equals(name, geomColumn, StringComparison.OrdinalIgnoreCase))
                        {
                            geomIndex = i;
                        }
                        else
                        {
                            nonGeomColumnNames.Add(name);
                            nonGeomIndices.Add(i);
                            nonGeomReaders.Add(TypeMapping.BuildColumnReader(reader.GetDataTypeName(i), typed));
                        }
                    }

                    if (geomIndex < 0)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                            $"Result has no column named '{geomColumn}'.");
                        return;
                    }

                    var geometries = new List<IGH_Goo>();
                    var dataTree = new GH_Structure<IGH_Goo>();
                    for (var branch = 0; branch < nonGeomColumnNames.Count; branch++)
                        dataTree.EnsurePath(new GH_Path(branch));

                    var rowCount = 0;
                    string warningAccumulator = null;

                    while (reader.Read())
                    {
                        var rawGeom = reader.IsDBNull(geomIndex) ? null : reader.GetValue(geomIndex);
                        geometries.Add(DecodeGeometry(rawGeom, ref warningAccumulator));

                        for (var b = 0; b < nonGeomIndices.Count; b++)
                        {
                            var colIdx = nonGeomIndices[b];
                            var goo = reader.IsDBNull(colIdx) ? null : nonGeomReaders[b](reader, colIdx);
                            dataTree.Append(goo, new GH_Path(b));
                        }
                        rowCount++;
                    }

                    if (!string.IsNullOrEmpty(warningAccumulator))
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, warningAccumulator);

                    da.SetDataList(_outGeometry, geometries);
                    da.SetDataList(_outColumns, nonGeomColumnNames);
                    da.SetDataTree(_outData, dataTree);
                    da.SetData(_outRows, rowCount);

                    if (limit > 0 && rowCount >= limit)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                            $"Row limit of {limit.ToString("N0", System.Globalization.CultureInfo.InvariantCulture)} reached; results may be truncated. " +
                            "Increase the Limit input (or set to 0 for no limit) to fetch more rows.");
                    }
                });
            }
            catch (Exception ex)
            {
                ReportError("Query failed", ex);
            }
        }


        private static IGH_Goo DecodeGeometry(object raw, ref string warningAccumulator)
        {
            if (raw is null) return null;

            byte[] bytes;
            switch (raw)
            {
                case byte[] b:
                    bytes = b;
                    break;
                case Stream stream:
                    using (var ms = new MemoryStream())
                    {
                        stream.CopyTo(ms);
                        bytes = ms.ToArray();
                    }
                    break;
                case ArraySegment<byte> seg:
                    bytes = new byte[seg.Count];
                    Array.Copy(seg.Array!, seg.Offset, bytes, 0, seg.Count);
                    break;
                default:
                    warningAccumulator = string.IsNullOrEmpty(warningAccumulator)
                        ? "Unexpected geometry value type: " + raw.GetType().Name
                        : warningAccumulator + " Unexpected geometry value type: " + raw.GetType().Name;
                    return null;
            }

            try
            {
                var geom = WkbCodec.Decode(bytes, out var w);
                if (!string.IsNullOrEmpty(w))
                    warningAccumulator = string.IsNullOrEmpty(warningAccumulator) ? w : warningAccumulator + " " + w;

                return WrapAsGoo(geom);
            }
            catch (Exception ex)
            {
                warningAccumulator = string.IsNullOrEmpty(warningAccumulator)
                    ? "Decode failure: " + ex.Message
                    : warningAccumulator + " Decode failure: " + ex.Message;
                return null;
            }
        }

        private static IGH_Goo WrapAsGoo(GeometryBase geom)
        {
            return geom switch
            {
                Point pt => new GH_Point(pt.Location),
                PointCloud cloud => new GH_PointCloud(cloud),
                Curve crv => new GH_Curve(crv),
                Mesh mesh => new GH_Mesh(mesh),
                Brep brep => new GH_Brep(brep),
                Surface srf => new GH_Surface(srf),
                _ => null
            };
        }
    }
}
