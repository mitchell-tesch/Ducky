using System;
using System.IO;
using GhDucky.Services;
using GhDucky.Utils;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;

namespace GhDucky.Components.Excel
{
    /// <summary>
    /// Writes a data tree into a cell or range of an xlsx/xlsm workbook,
    /// without Excel or COM.
    /// </summary>
    public class ExcelWriteComponent : DuckyComponentBase
    {
        public ExcelWriteComponent()
            : base(
                "Ducky Excel Write",
                "DuckyXlWrite",
                "Writes a data tree into an Excel workbook (.xlsx/.xlsm). " +
                "Does not use Excel or COM. In place mode edits only the target cells, leaving " +
                "all other cells, their formatting, the sheet's formulas and other sheets intact.",
                "Ducky",
                "5 | Excel")
        {
        }

        public override Guid ComponentGuid => new Guid("b7d5e441-2c96-4a18-8f77-1e3c6d84b522");

        public override GH_Exposure Exposure => GH_Exposure.primary;

        protected override System.Drawing.Bitmap Icon => IconFactory.Build("✏️", IconFactory.Excel);

        private int _inWrite;
        private int _inPath;
        private int _inSheet;
        private int _inRange;
        private int _inValues;
        private int _inDirection;
        private int _inMode;
        private int _inFormulas;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            _inWrite = pManager.AddBooleanParameter("Write?", "W?",
                "Set to true to perform the write.",
                GH_ParamAccess.item, false);
            _inPath = pManager.AddTextParameter("Path", "P",
                "Absolute path to the workbook (.xlsx or .xlsm). Created if it does not exist.",
                GH_ParamAccess.item);
            _inSheet = pManager.AddTextParameter("Sheet", "S",
                "Worksheet name. Leave empty to use the first sheet. Created if it does not exist.",
                GH_ParamAccess.item, string.Empty);
            _inRange = pManager.AddTextParameter("Range", "Rg",
                "Anchor cell or range in A1 notation (e.g. 'B2' or 'B2:D10'), optionally " +
                "sheet-qualified. An anchor alone auto-expands to fit the data; an explicit " +
                "range clips it.",
                GH_ParamAccess.item, "A1");
            _inValues = pManager.AddGenericParameter("Values", "V",
                "Values to write, as a data tree. Numbers, dates, booleans and text keep their " +
                "type. Null or empty items clear the target cell.",
                GH_ParamAccess.tree);
            _inDirection = pManager.AddIntegerParameter("Direction", "D",
                "Tree orientation. 0 = By Row (each branch is a spreadsheet row), " +
                "1 = By Column (each branch is a spreadsheet column).",
                GH_ParamAccess.item, 0);
            _inMode = pManager.AddIntegerParameter("Mode", "M",
                "0 = In Place (edit only the target cells, preserve everything else), " +
                "1 = Overwrite Sheet (clear the sheet first, keep other sheets), " +
                "2 = New Workbook (replace the file with a fresh single-sheet workbook).",
                GH_ParamAccess.item, 0);
            _inFormulas = pManager.AddBooleanParameter("Formulas?", "F?",
                "If true, text beginning with '=' is written as a formula rather than literal text.",
                GH_ParamAccess.item, true);

            if (pManager[_inDirection] is Param_Integer direction)
            {
                direction.AddNamedValue("By Row", 0);
                direction.AddNamedValue("By Column", 1);
            }

            if (pManager[_inMode] is Param_Integer mode)
            {
                mode.AddNamedValue("In Place", 0);
                mode.AddNamedValue("Overwrite Sheet", 1);
                mode.AddNamedValue("New Workbook", 2);
            }

            pManager[_inSheet].Optional = true;
            pManager[_inRange].Optional = true;
            pManager[_inDirection].Optional = true;
            pManager[_inMode].Optional = true;
            pManager[_inFormulas].Optional = true;
        }

        private int _outPath;
        private int _outAddress;
        private int _outCells;

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            _outPath = pManager.AddTextParameter("Path", "P",
                "Path to the written workbook (passthrough).",
                GH_ParamAccess.item);
            _outAddress = pManager.AddTextParameter("Address", "A",
                "The address actually written, e.g. 'Data!B2:D11'.",
                GH_ParamAccess.item);
            _outCells = pManager.AddIntegerParameter("Cells", "C",
                "Number of cells written.",
                GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            var write = false;
            da.GetData(_inWrite, ref write);
            if (!write)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Write is false; no action taken.");
                return;
            }

            var path = string.Empty;
            if (!da.GetData(_inPath, ref path) || string.IsNullOrWhiteSpace(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Path to workbook required.");
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

            if (!da.GetDataTree(_inValues, out GH_Structure<IGH_Goo> tree) || tree is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "No values supplied.");
                return;
            }

            var sheet = string.Empty;
            var range = "A1";
            var direction = 0;
            var mode = 0;
            var formulas = true;
            da.GetData(_inSheet, ref sheet);
            da.GetData(_inRange, ref range);
            da.GetData(_inDirection, ref direction);
            da.GetData(_inMode, ref mode);
            da.GetData(_inFormulas, ref formulas);

            if (string.IsNullOrWhiteSpace(range)) range = "A1";

            try
            {
                if (!TryResolveAddress(path, range, sheet, out var address, out var addressError))
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, addressError);
                    return;
                }

                var grid = BuildGrid(tree, direction == 1);
                if (grid.Length == 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "No values to write.");
                    return;
                }

                var result = ExcelWriter.Write(path, address, grid, (ExcelWriteMode)mode, formulas);

                if (result.Clipped)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                        "The data was larger than the given range and has been clipped to " +
                        result.Address + ".");
                }

                da.SetData(_outPath, Path.GetFullPath(path));
                da.SetData(_outAddress, result.Address);
                da.SetData(_outCells, result.CellsWritten);
            }
            catch (Exception ex)
            {
                ReportError("Excel write failed", ex);
            }
        }

        /// <summary>
        /// Parses the range and resolves defined names / Table names against the
        /// workbook when it already exists.
        /// </summary>
        private static bool TryResolveAddress(
            string path, string range, string sheet, out ExcelAddress address, out string error)
        {
            if (!ExcelAddress.TryParse(range, out address, out error))
                return false;

            if (address.IsName)
            {
                if (!File.Exists(path))
                {
                    error = $"'{address.Name}' is a name, but the workbook does not exist yet so it cannot be resolved.";
                    return false;
                }

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

        /// <summary>
        /// Flattens the tree into a row-major grid in sheet orientation.
        /// By Row treats each branch as a spreadsheet row; By Column transposes,
        /// so each branch becomes a spreadsheet column.
        /// </summary>
        private static object[][] BuildGrid(GH_Structure<IGH_Goo> tree, bool byColumn)
        {
            var branches = tree.Branches;
            if (branches.Count == 0) return Array.Empty<object[]>();

            var longest = 0;
            foreach (var branch in branches)
            {
                if (branch != null && branch.Count > longest) longest = branch.Count;
            }
            if (longest == 0) return Array.Empty<object[]>();

            if (!byColumn)
            {
                // Each branch is a row; ragged branches are padded with nulls,
                // which clear rather than overwrite the trailing cells.
                var grid = new object[branches.Count][];
                for (var r = 0; r < branches.Count; r++)
                {
                    var branch = branches[r];
                    var row = new object[longest];
                    if (branch != null)
                    {
                        for (var c = 0; c < branch.Count; c++)
                            row[c] = Unwrap(branch[c]);
                    }
                    grid[r] = row;
                }
                return grid;
            }

            // Each branch is a column.
            var transposed = new object[longest][];
            for (var r = 0; r < longest; r++)
            {
                var row = new object[branches.Count];
                for (var c = 0; c < branches.Count; c++)
                {
                    var branch = branches[c];
                    row[c] = branch != null && r < branch.Count ? Unwrap(branch[r]) : null;
                }
                transposed[r] = row;
            }
            return transposed;
        }

        /// <summary>
        /// Unwraps goo to a CLR value. Empty strings are normalised to null so
        /// they clear the cell rather than writing a zero-length string.
        /// </summary>
        private static object Unwrap(IGH_Goo goo)
        {
            var value = TypeMapping.Unwrap(goo);
            if (value is string s && s.Length == 0) return null;
            return value;
        }
    }
}



