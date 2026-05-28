using System;
using System.IO;
using GhDucky.Parameters;
using GhDucky.Utils;
using Grasshopper.Kernel;

namespace GhDucky.Components.Import
{
    public class ImportJsonComponent : DuckyComponentBase
    {
        public ImportJsonComponent()
            : base(
                "Ducky Import JSON",
                "DuckyJSON",
                "Loads a JSON file into a Database table (using DuckDB's native read_json_auto reader). " +
                "Supports both line-delimited and array-style JSON files.",
                "Ducky",
                "2 | Import")
        {
        }

        public override Guid ComponentGuid => new Guid("a86d8d62-50bd-4e7b-8f8c-7d2cb3a45e22");

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        protected override System.Drawing.Bitmap Icon => IconFactory.Build("🦢", IconFactory.ImportFile);

        private int _inImport;
        private int _inDatabase;
        private int _inPath;
        private int _inTable;
        private int _inSchema;
        private int _inFormat;
        private int _inOverwrite;
        private int _inSampleSize;
        private int _inMaxObjectSize;

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
                "Absolute path to source JSON file.",
                GH_ParamAccess.item);
            _inFormat = pManager.AddTextParameter("Format", "F",
                "JSON format: auto (default), array, newline_delimited.",
                GH_ParamAccess.item, "auto");
            _inSchema = pManager.AddTextParameter("Schema", "Sc",
                "Target schema (default: main). Created automatically if missing.",
                GH_ParamAccess.item, "main");
            _inOverwrite = pManager.AddBooleanParameter("Overwrite?", "O?",
                "If true (default), the table is dropped and recreated; otherwise rows are appended.",
                GH_ParamAccess.item, true);
            _inSampleSize = pManager.AddIntegerParameter("Sample Size", "SS",
                "Number of objects sampled for type inference. " +
                "Default 20480 is usually sufficient. " +
                "Use -1 to scan the entire file (slow for large files).",
                GH_ParamAccess.item, 20480);
            _inMaxObjectSize = pManager.AddIntegerParameter("Max Object Size", "MO",
                "Maximum size (in bytes) of a single JSON object. " +
                "Increase beyond the default 16777216 (16 MB) for files with very large nested objects.",
                GH_ParamAccess.item, 16777216);

            pManager[_inFormat].Optional = true;
            pManager[_inSchema].Optional = true;
            pManager[_inSampleSize].Optional = true;
            pManager[_inMaxObjectSize].Optional = true;
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

            string table = null;
            if (!da.GetData(_inTable, ref table) || string.IsNullOrWhiteSpace(table))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Table name required.");
                return;
            }

            string path = null;
            if (!da.GetData(_inPath, ref path) || string.IsNullOrWhiteSpace(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Path to JSON file required.");
                return;
            }
            if (!File.Exists(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "JSON file not found: " + path);
                return;
            }

            var format = "auto";
            da.GetData(_inFormat, ref format);
            var normalizedFormat = (format ?? "auto").Trim().ToLowerInvariant();
            if (normalizedFormat != "auto" &&
                normalizedFormat != "array" &&
                normalizedFormat != "newline_delimited")
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                    $"Unknown format '{format}', falling back to 'auto'.");
                normalizedFormat = "auto";
            }

            var schema = "main";
            var overwrite = true;
            var sampleSize = 20480;
            var maxObjectSize = 16777216;
            da.GetData(_inSchema, ref schema);
            da.GetData(_inOverwrite, ref overwrite);
            da.GetData(_inSampleSize, ref sampleSize);
            da.GetData(_inMaxObjectSize, ref maxObjectSize);

            var quotedTable = SqlIdentifier.QuoteTable(schema, table);
            var pathLiteral = SqlIdentifier.QuoteLiteral(Path.GetFullPath(path));
            var formatLiteral = SqlIdentifier.QuoteLiteral(normalizedFormat);
            var sourceExpr = $"read_json({pathLiteral}, format={formatLiteral}, sample_size={sampleSize}, maximum_object_size={maxObjectSize})";

            try
            {
                session.Execute(conn =>
                {
                    var rowCount = RunSelectImport(conn, schema, table, quotedTable, sourceExpr, overwrite);

                    da.SetData(_outDatabase, dbConnection);
                    da.SetData(_outTable, table);
                    da.SetData(_outRows, rowCount);
                });
            }
            catch (Exception ex)
            {
                ReportError("JSON import failed", ex);
            }
        }
    }
}
