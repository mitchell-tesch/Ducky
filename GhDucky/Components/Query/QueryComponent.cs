using System;
using System.Collections.Generic;
using System.Globalization;
using DuckDB.NET.Data;
using GhDucky.Parameters;
using GhDucky.Utils;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;

namespace GhDucky.Components.Query
{
    public class QueryComponent : DuckyComponentBase
    {
        public QueryComponent()
            : base(
                "Ducky Query",
                "DuckyQ",
                "Executes a SQL query against the DuckDB connection and returns the result as a Grasshopper data tree. " +
                "Each branch is one column; items within a branch are row values, ordered as returned by the query.",
                "Ducky",
                "3 | Query")
        {
        }

        public override Guid ComponentGuid => new Guid("c2b6e21a-7e4d-4d70-9b6f-3a1c5a2a59e9");

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        protected override System.Drawing.Bitmap Icon => IconFactory.Build("❓", IconFactory.Query);

        private int _inRun;
        private int _inDatabase;
        private int _inQuery;
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
            _inQuery = pManager.AddTextParameter("Query", "Q",
                "SQL query to execute. SELECT statements return a result tree; other statements report the rows-affected count.",
                GH_ParamAccess.item);
            _inLimit = pManager.AddIntegerParameter("Limit", "L",
                "Maximum number of rows to fetch. Defaults to 1,000,000 to guard against excessive memory use; " +
                "set to 0 to return all rows (use with care on large result sets).",
                GH_ParamAccess.item, DefaultRowLimit);
            _inTyped = pManager.AddBooleanParameter("Typed?", "T?",
                "If true (default), temporal columns (DATE/TIME/TIMESTAMP) surface as GH_Time so downstream date arithmetic works. " +
                "If false, those values are emitted as ISO 8601 strings.",
                GH_ParamAccess.item, true);

            pManager[_inLimit].Optional = true;
            pManager[_inTyped].Optional = true;
        }

        private int _outColumns;
        private int _outTypes;
        private int _outData;
        private int _outRows;

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            _outColumns = pManager.AddTextParameter("Columns", "N",
                "Column names, parallel to the branches of Data.",
                GH_ParamAccess.list);
            _outTypes = pManager.AddTextParameter("Types", "T",
                "Column data types, parallel to the branches of Data.",
                GH_ParamAccess.list);
            _outData = pManager.AddGenericParameter("Data", "D",
                "Result data tree: one branch per column, items are row values.",
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

            var query = string.Empty;
            if (!da.GetData(_inQuery, ref query) || string.IsNullOrWhiteSpace(query))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "SQL query required.");
                return;
            }

            var limit = DefaultRowLimit;
            da.GetData(_inLimit, ref limit);
            if (limit < 0) limit = 0;

            var typed = true;
            da.GetData(_inTyped, ref typed);

            try
            {
                if (IsLikelyResultProducing(query))
                {
                    var effectiveSql = limit > 0
                        ? $"SELECT * FROM ({query.TrimEnd(';')}) AS __ghducky_sub LIMIT {limit.ToString(CultureInfo.InvariantCulture)}"
                        : query;
                    session.Execute(conn => ExecuteReader(conn, effectiveSql, typed, limit, da));
                }
                else
                {
                    session.Execute(conn =>
                    {
                        using var cmd = conn.CreateCommand();
                        cmd.CommandText = query;
                        var affected = cmd.ExecuteNonQuery();

                        da.SetDataList(_outColumns, new List<string>());
                        da.SetDataList(_outTypes, new List<string>());
                        da.SetDataTree(_outData, new GH_Structure<IGH_Goo>());
                        da.SetData(_outRows, affected);
                    });
                }
            }
            catch (Exception ex)
            {
                ReportError("Query failed", ex);
            }
        }


        private static bool IsLikelyResultProducing(string sql)
        {
            var trimmed = sql.TrimStart();

            // Strip leading comments (single-line and block).
            while (true)
            {
                if (trimmed.StartsWith("--", StringComparison.Ordinal))
                {
                    var nl = trimmed.IndexOf('\n');
                    if (nl < 0) return false;
                    trimmed = trimmed.Substring(nl + 1).TrimStart();
                }
                else if (trimmed.StartsWith("/*", StringComparison.Ordinal))
                {
                    var end = trimmed.IndexOf("*/", 2, StringComparison.Ordinal);
                    if (end < 0) return false;
                    trimmed = trimmed.Substring(end + 2).TrimStart();
                }
                else
                {
                    break;
                }
            }

            return trimmed.StartsWith("SELECT", StringComparison.OrdinalIgnoreCase) ||
                   trimmed.StartsWith("WITH", StringComparison.OrdinalIgnoreCase) ||
                   trimmed.StartsWith("SHOW", StringComparison.OrdinalIgnoreCase) ||
                   trimmed.StartsWith("DESCRIBE", StringComparison.OrdinalIgnoreCase) ||
                   trimmed.StartsWith("PRAGMA", StringComparison.OrdinalIgnoreCase) ||
                   trimmed.StartsWith("EXPLAIN", StringComparison.OrdinalIgnoreCase) ||
                   trimmed.StartsWith("VALUES", StringComparison.OrdinalIgnoreCase) ||
                   trimmed.StartsWith("TABLE", StringComparison.OrdinalIgnoreCase) ||
                   trimmed.StartsWith("FROM", StringComparison.OrdinalIgnoreCase);
        }


        private void ExecuteReader(DuckDBConnection dbConnection, string sql, bool typed, int limit, IGH_DataAccess da)
        {
            using var cmd = dbConnection.CreateCommand();
            cmd.CommandText = sql;

            using var reader = cmd.ExecuteReader();

            var columnCount = reader.FieldCount;
            var columnNames = new List<string>(columnCount);
            var columnTypes = new List<string>(columnCount);
            var paths = new GH_Path[columnCount];
            var readers = new Func<System.Data.Common.DbDataReader, int, IGH_Goo>[columnCount];

            for (var i = 0; i < columnCount; i++)
            {
                var typeName = reader.GetDataTypeName(i);
                columnNames.Add(reader.GetName(i));
                columnTypes.Add(typeName);
                paths[i] = new GH_Path(i);
                readers[i] = TypeMapping.BuildColumnReader(typeName, typed);
            }

            var dataTree = new GH_Structure<IGH_Goo>();
            for (var i = 0; i < columnCount; i++)
                dataTree.EnsurePath(paths[i]);

            var rowCount = 0;
            while (reader.Read())
            {
                for (var i = 0; i < columnCount; i++)
                {
                    var goo = reader.IsDBNull(i) ? null : readers[i](reader, i);
                    dataTree.Append(goo, paths[i]);
                }
                rowCount++;
            }

            da.SetDataList(_outColumns, columnNames);
            da.SetDataList(_outTypes, columnTypes);
            da.SetDataTree(_outData, dataTree);
            da.SetData(_outRows, rowCount);

            if (limit > 0 && rowCount >= limit)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                    $"Row limit of {limit.ToString("N0", CultureInfo.InvariantCulture)} reached; results may be truncated. " +
                    "Increase the Limit input (or set to 0 for no limit) to fetch more rows.");
            }
        }
    }
}
