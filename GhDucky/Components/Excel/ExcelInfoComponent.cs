using System;
using System.IO;
using System.Linq;
using GhDucky.Services;
using GhDucky.Utils;
using Grasshopper.Kernel;

namespace GhDucky.Components.Excel
{
    /// <summary>
    /// Inspects an xlsx/xlsm workbook: sheets, used ranges, defined names and
    /// Excel Tables, plus basic file state.
    /// </summary>
    public class ExcelInfoComponent : DuckyComponentBase
    {
        public ExcelInfoComponent()
            : base(
                "Ducky Excel Info",
                "DuckyXlInfo",
                "Inspects an Excel workbook (.xlsx/.xlsm): sheet names, used ranges, dimensions, " +
                "hidden state, defined names (named ranges), Excel Tables and file state. " +
                "Useful for driving the Range input of Ducky Excel Read/Write.",
                "Ducky",
                "5 | Excel")
        {
        }

        public override Guid ComponentGuid => new Guid("9c86f2ad-3b51-4d70-b0a4-5f9e21c7d833");

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        protected override System.Drawing.Bitmap Icon => IconFactory.Build("🔎", IconFactory.Excel);

        private int _inPath;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            _inPath = pManager.AddTextParameter("Path", "P",
                "Absolute path to the workbook (.xlsx or .xlsm).",
                GH_ParamAccess.item);
        }

        private int _outSheets;
        private int _outUsedRanges;
        private int _outRows;
        private int _outColumns;
        private int _outHidden;
        private int _outNames;
        private int _outNameRefs;
        private int _outTables;
        private int _outTableRefs;
        private int _outSize;
        private int _outModified;
        private int _outLocked;

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            _outSheets = pManager.AddTextParameter("Sheets", "S",
                "Sheet names, in workbook order.",
                GH_ParamAccess.list);
            _outUsedRanges = pManager.AddTextParameter("Used Range", "UR",
                "Used range of each sheet (e.g. 'A1:F204'), parallel to Sheets.",
                GH_ParamAccess.list);
            _outRows = pManager.AddIntegerParameter("Rows", "R",
                "Row count of each sheet's used range, parallel to Sheets.",
                GH_ParamAccess.list);
            _outColumns = pManager.AddIntegerParameter("Columns", "C",
                "Column count of each sheet's used range, parallel to Sheets.",
                GH_ParamAccess.list);
            _outHidden = pManager.AddBooleanParameter("Hidden", "H",
                "Whether each sheet is hidden, parallel to Sheets.",
                GH_ParamAccess.list);
            _outNames = pManager.AddTextParameter("Named Ranges", "N",
                "Defined names in the workbook (Excel built-ins excluded).",
                GH_ParamAccess.list);
            _outNameRefs = pManager.AddTextParameter("Name Refs", "NR",
                "Reference each defined name points at, parallel to Named Ranges.",
                GH_ParamAccess.list);
            _outTables = pManager.AddTextParameter("Tables", "T",
                "Excel Table names.",
                GH_ParamAccess.list);
            _outTableRefs = pManager.AddTextParameter("Table Refs", "TR",
                "Sheet-qualified extent of each Table, parallel to Tables.",
                GH_ParamAccess.list);
            _outSize = pManager.AddIntegerParameter("Size", "Sz",
                "File size in bytes.",
                GH_ParamAccess.item);
            _outModified = pManager.AddTimeParameter("Modified", "M",
                "Last write time of the file.",
                GH_ParamAccess.item);
            _outLocked = pManager.AddBooleanParameter("Locked?", "L?",
                "True when the workbook is currently locked for writing (typically open in Excel).",
                GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
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

            try
            {
                var sheets = ExcelReader.GetSheets(path);
                var names = ExcelWorkbookMetadata.ReadDefinedNames(path);
                var tables = ExcelWorkbookMetadata.ReadTables(path);
                var info = new FileInfo(Path.GetFullPath(path));

                da.SetDataList(_outSheets, sheets.Select(s => s.Name).ToList());
                da.SetDataList(_outUsedRanges, sheets.Select(s => s.UsedRange).ToList());
                da.SetDataList(_outRows, sheets.Select(s => s.RowCount).ToList());
                da.SetDataList(_outColumns, sheets.Select(s => s.ColumnCount).ToList());
                da.SetDataList(_outHidden, sheets.Select(s => s.Hidden).ToList());

                da.SetDataList(_outNames, names.Select(n => n.Name).ToList());
                da.SetDataList(_outNameRefs, names.Select(n => n.Reference).ToList());

                da.SetDataList(_outTables, tables.Select(t => t.Name).ToList());
                da.SetDataList(_outTableRefs, tables
                    .Select(t => string.IsNullOrEmpty(t.Sheet) ? t.Reference : t.Sheet + "!" + t.Reference)
                    .ToList());

                da.SetData(_outSize, (int)Math.Min(info.Length, int.MaxValue));
                da.SetData(_outModified, info.LastWriteTime);
                da.SetData(_outLocked, ExcelFileGuard.IsLockedForWriting(path, out _));
            }
            catch (Exception ex)
            {
                ReportError("Excel inspection failed", ex);
            }
        }
    }
}


