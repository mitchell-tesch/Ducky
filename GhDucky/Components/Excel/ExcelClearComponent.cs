using System;
using System.IO;
using GhDucky.Services;
using GhDucky.Utils;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;

namespace GhDucky.Components.Excel
{
    /// <summary>
    /// Clears cells, ranges or whole sheets in an xlsx/xlsm workbook.
    /// </summary>
    public class ExcelClearComponent : DuckyComponentBase
    {
        public ExcelClearComponent()
            : base(
                "Ducky Excel Clear",
                "DuckyXlClear",
                "Clears a cell, a range, or an entire sheet in an Excel workbook (.xlsx/.xlsm). " +
                "Values only keeps the cells' formatting; Values and Formatting removes the cells " +
                "outright. Everything outside the target range is left untouched.",
                "Ducky",
                "5 | Excel")
        {
        }

        public override Guid ComponentGuid => new Guid("2a94b7e3-6d18-4c52-9b61-08f4a7c3e6d9");

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        protected override System.Drawing.Bitmap Icon => IconFactory.Build("🧽", IconFactory.Excel);

        private int _inClear;
        private int _inPath;
        private int _inSheet;
        private int _inRange;
        private int _inMode;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            _inClear = pManager.AddBooleanParameter("Clear?", "C?",
                "Set to true to perform the clear.",
                GH_ParamAccess.item, false);
            _inPath = pManager.AddTextParameter("Path", "P",
                "Absolute path to the workbook (.xlsx or .xlsm).",
                GH_ParamAccess.item);
            _inSheet = pManager.AddTextParameter("Sheet", "S",
                "Worksheet name. Leave empty to use the first sheet.",
                GH_ParamAccess.item, string.Empty);
            _inRange = pManager.AddTextParameter("Range", "Rg",
                "Cell or range to clear, in A1 notation, optionally sheet-qualified. " +
                "A single cell clears only that cell — clearing never auto-expands. " +
                "Ignored when Mode is Entire Sheet.",
                GH_ParamAccess.item, "A1");
            _inMode = pManager.AddIntegerParameter("Mode", "M",
                "0 = Values (clear contents, keep formatting), " +
                "1 = Values and Formatting (remove the cells entirely), " +
                "2 = Entire Sheet (clear every cell on the sheet; Range is ignored).",
                GH_ParamAccess.item, 0);

            if (pManager[_inMode] is Param_Integer mode)
            {
                mode.AddNamedValue("Values", 0);
                mode.AddNamedValue("Values and Formatting", 1);
                mode.AddNamedValue("Entire Sheet", 2);
            }

            pManager[_inSheet].Optional = true;
            pManager[_inRange].Optional = true;
            pManager[_inMode].Optional = true;
        }

        private int _outPath;
        private int _outAddress;
        private int _outCells;

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            _outPath = pManager.AddTextParameter("Path", "P",
                "Path to the workbook (passthrough).",
                GH_ParamAccess.item);
            _outAddress = pManager.AddTextParameter("Address", "A",
                "The address actually cleared, or the sheet name in Entire Sheet mode.",
                GH_ParamAccess.item);
            _outCells = pManager.AddIntegerParameter("Cells", "C",
                "Number of cells cleared.",
                GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            var clear = false;
            da.GetData(_inClear, ref clear);
            if (!clear)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Clear is false; no action taken.");
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

            if (ExcelFileGuard.IsLockedForWriting(path, out var lockReason))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, lockReason);
                return;
            }

            var sheet = string.Empty;
            var range = "A1";
            var mode = 0;
            da.GetData(_inSheet, ref sheet);
            da.GetData(_inRange, ref range);
            da.GetData(_inMode, ref mode);

            if (string.IsNullOrWhiteSpace(range)) range = "A1";

            try
            {
                if (!ExcelAddress.TryParse(range, out var address, out var addressError))
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, addressError);
                    return;
                }

                if (address.IsName)
                {
                    var reference = ExcelWorkbookMetadata.ResolveName(path, address.Name,
                        string.IsNullOrWhiteSpace(sheet) ? null : sheet);

                    if (reference is null)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                            $"'{address.Name}' is not a defined name or Table in this workbook.");
                        return;
                    }

                    if (!ExcelAddress.TryParse(reference, out address, out addressError))
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error,
                            $"Defined name '{range}' resolves to '{reference}', which could not be parsed.");
                        return;
                    }
                }

                if (address.Sheet is null && !string.IsNullOrWhiteSpace(sheet))
                    address = address.WithSheet(sheet);

                var result = ExcelWriter.Clear(path, address, (ExcelClearMode)mode);

                da.SetData(_outPath, Path.GetFullPath(path));
                da.SetData(_outAddress, result.Address);
                da.SetData(_outCells, result.CellsWritten);

                if (result.CellsWritten == 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark,
                        "No populated cells were found in the target range.");
                }
            }
            catch (Exception ex)
            {
                ReportError("Excel clear failed", ex);
            }
        }
    }
}

