using System;
using System.Collections.Generic;
using System.IO;
using GhDucky.Parameters;
using GhDucky.Utils;
using Grasshopper.Kernel;

namespace GhDucky.Components.Export
{
    public class ExportComponent : DuckyComponentBase
    {
        public ExportComponent()
            : base(
                "Ducky Export",
                "DuckyEx",
                "Exports a table or query result to disk via DuckDB's COPY TO statement. " +
                "The Source input accepts either a table name or a full SELECT statement.",
                "Ducky",
                "4 | Export")
        {
        }

        public override Guid ComponentGuid => new Guid("2f3a6f0b-1d5c-4b9b-9a3e-6d2c1e8a7f31");

        protected override System.Drawing.Bitmap Icon => IconFactory.Build("📦", IconFactory.Export);

        public override GH_Exposure Exposure => GH_Exposure.primary;

        private int _inExport;
        private int _inDatabase;
        private int _inSource;
        private int _inPath;
        private int _inFormat;
        private int _inIncludeHeader;
        private int _inOverwrite;
        private int _inDelimiter;
        private int _inCompression;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            _inExport = pManager.AddBooleanParameter("Export?", "E?",
                "Set to true to perform the export.",
                GH_ParamAccess.item, false);
            _inDatabase = pManager.AddParameter(new ParamDuckyDbConnection(), "Database", "DB",
                "Database connection.",
                GH_ParamAccess.item);
            _inSource = pManager.AddTextParameter("Source", "S",
                "Table name OR a SELECT statement (anything that can be wrapped in COPY (...) TO).",
                GH_ParamAccess.item);
            _inPath = pManager.AddTextParameter("Path", "P",
                "Absolute output file path.",
                GH_ParamAccess.item);
            _inFormat = pManager.AddTextParameter("Format", "F",
                "Output format: CSV, JSON, or PARQUET. Defaults to inferring from the path extension.",
                GH_ParamAccess.item, string.Empty);
            _inIncludeHeader = pManager.AddBooleanParameter("Header?", "H?",
                "Include header row (for CSV only). Ignored for other formats.",
                GH_ParamAccess.item, true);
            _inOverwrite = pManager.AddBooleanParameter("Overwrite?", "O?",
                "If true, an existing output file is overwritten.",
                GH_ParamAccess.item, true);
            _inDelimiter = pManager.AddTextParameter("Delimiter", "D",
                "Column delimiter for CSV export (e.g. ',' or '\\t'). " +
                "Defaults to comma. Ignored for other formats.",
                GH_ParamAccess.item, string.Empty);
            _inCompression = pManager.AddTextParameter("Compression", "Cmp",
                "Compression codec for Parquet export: SNAPPY (default, fast), " +
                "ZSTD (smaller files), GZIP, LZ4, or UNCOMPRESSED (fastest write). Ignored for other formats.",
                GH_ParamAccess.item, string.Empty);

            pManager[_inFormat].Optional = true;
            pManager[_inIncludeHeader].Optional = true;
            pManager[_inDelimiter].Optional = true;
            pManager[_inCompression].Optional = true;
        }

        private int _outDatabase;
        private int _outPath;

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            _outDatabase = pManager.AddParameter(new ParamDuckyDbConnection(), "Database", "DB",
                "Database connection (passthrough).",
                GH_ParamAccess.item);
            _outPath = pManager.AddTextParameter("Path", "P",
                "Output file path.",
                GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            var export = false;
            da.GetData(_inExport, ref export);
            if (!export)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Export is false; no action taken.");
                return;
            }

            if (!TryGetSession(da, _inDatabase, out var session, out var dbConnection))
                return;

            var source = string.Empty;
            if (!da.GetData(_inSource, ref source) || string.IsNullOrWhiteSpace(source))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Source table name or SELECT statement required.");
                return;
            }

            var path = string.Empty;
            if (!da.GetData(_inPath, ref path) || string.IsNullOrWhiteSpace(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Path for export file required.");
                return;
            }

            var overwrite = true;
            da.GetData(_inOverwrite, ref overwrite);

            var absolutePath = Path.GetFullPath(path);
            if (File.Exists(absolutePath))
            {
                if (!overwrite)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                        "Output file already exists and Overwrite is false: " + absolutePath);
                    return;
                }
                try { File.Delete(absolutePath); }
                catch (Exception ex)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                        "Could not remove existing output file: " + ex.Message);
                    return;
                }
            }

            var dir = Path.GetDirectoryName(absolutePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                try { Directory.CreateDirectory(dir); }
                catch (Exception ex)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                        "Could not create output directory: " + ex.Message);
                    return;
                }
            }

            var format = string.Empty;
            da.GetData(_inFormat, ref format);
            var resolvedFormat = ResolveFormat(format, path);
            if (resolvedFormat == null)
            {
                var ext = (Path.GetExtension(path) ?? string.Empty).ToLowerInvariant();
                var hint = ext is ".xlsx" or ".xls"
                    ? " For Excel files, use the Ducky Export Excel component."
                    : string.Empty;
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                    "Could not determine output format. Provide Format=CSV|JSON|PARQUET or a recognised file extension." + hint);
                return;
            }

            var header = true;
            da.GetData(_inIncludeHeader, ref header);

            var delimiter = string.Empty;
            var compression = string.Empty;
            da.GetData(_inDelimiter, ref delimiter);
            da.GetData(_inCompression, ref compression);

            var sourceExpression = WrapSource(source);
            var optionsClause = BuildOptionsClause(resolvedFormat, header, delimiter, compression);

            try
            {
                session.Execute(conn =>
                {
                    using var cmd = conn.CreateCommand();
                    cmd.CommandText =
                        $"COPY {sourceExpression} TO {SqlIdentifier.QuoteLiteral(absolutePath)} {optionsClause}";
                    cmd.ExecuteNonQuery();
                });

                da.SetData(_outDatabase, dbConnection);
                da.SetData(_outPath, path);
            }
            catch (Exception ex)
            {
                ReportError("Export failed", ex);
            }
        }


        private static string ResolveFormat(string requested, string path)
        {
            if (!string.IsNullOrWhiteSpace(requested))
            {
                var f = requested.Trim().ToUpperInvariant();
                return f is "CSV" or "JSON" or "PARQUET" ? f : null;
            }

            var ext = (Path.GetExtension(path) ?? string.Empty).ToLowerInvariant();
            return ext switch
            {
                ".csv" or ".tsv" => "CSV",
                ".json" or ".ndjson" or ".jsonl" => "JSON",
                ".parquet" or ".pq" => "PARQUET",
                _ => null
            };
        }



        private static string BuildOptionsClause(string format, bool header, string delimiter, string compression)
        {
            var opts = new List<string> { "FORMAT '" + format + "'" };

            if (format == "CSV")
            {
                opts.Add("HEADER " + (header ? "TRUE" : "FALSE"));
                if (!string.IsNullOrWhiteSpace(delimiter))
                    opts.Add("DELIMITER " + SqlIdentifier.QuoteLiteral(delimiter));
            }

            if (format == "PARQUET" && !string.IsNullOrWhiteSpace(compression))
            {
                var codec = compression.Trim().ToUpperInvariant();
                // Normalize common aliases to the DuckDB-recognized name.
                if (codec is "NONE") codec = "UNCOMPRESSED";
                if (codec is "SNAPPY" or "ZSTD" or "GZIP" or "LZ4" or "UNCOMPRESSED")
                    opts.Add("CODEC '" + codec + "'");
            }

            return "(" + string.Join(", ", opts) + ")";
        }
    }
}
