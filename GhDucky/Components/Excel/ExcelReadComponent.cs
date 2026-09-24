using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using GhDucky.Services;
using GhDucky.Utils;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;

namespace GhDucky.Components.Excel
{
    /// <summary>
    /// Reads a cell or range from an xlsx/xlsm workbook into a data tree,
    /// without Excel or COM.
    /// </summary>
    public class ExcelReadComponent : DuckyComponentBase
    {
        public ExcelReadComponent()
            : base(
                "Ducky Excel Read",
                "DuckyXlRead",
                "Reads a cell or range from an Excel workbook (.xlsx/.xlsm) into a data tree. " +
                "Does not use Excel or COM, so it is fast and works headlessly. " +
                "Give a single cell with Expand? enabled to read the whole contiguous block around it.",
                "Ducky",
                "5 | Excel")
        {
        }

        public override Guid ComponentGuid => new Guid("4f3a1c62-58b0-4e9d-9d3f-7b2a5e6c9a11");

        public override GH_Exposure Exposure => GH_Exposure.primary;

        protected override System.Drawing.Bitmap Icon => IconFactory.Build("📖", IconFactory.Excel);

        private int _inRead;
        private int _inPath;
        private int _inSheet;
        private int _inRange;
        private int _inExpand;
        private int _inHeaders;
        private int _inDirection;
        private int _inFormulas;
        private int _inCache;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            _inRead = pManager.AddBooleanParameter("Read?", "R?",
                "Set to true to perform the read.",
                GH_ParamAccess.item, true);
            _inPath = pManager.AddTextParameter("Path", "P",
                "Absolute path to the workbook (.xlsx or .xlsm).",
                GH_ParamAccess.item);
            _inSheet = pManager.AddTextParameter("Sheet", "S",
                "Worksheet name. Leave empty to use the first sheet.",
                GH_ParamAccess.item, string.Empty);
            _inRange = pManager.AddTextParameter("Range", "Rg",
                "Cell or range in A1 notation (e.g. 'B2' or 'B2:D10'), optionally sheet-qualified " +
                "('Data!B2:D10'). A defined name or Excel Table name may also be given.",
                GH_ParamAccess.item, "A1");
            _inExpand = pManager.AddBooleanParameter("Expand?", "E?",
                "When a single anchor cell is given, expand to the enclosing contiguous block " +
                "(a fully blank row bounds it vertically, a fully blank column horizontally). " +
                "Ignored when an explicit range is supplied.",
                GH_ParamAccess.item, true);
            _inHeaders = pManager.AddBooleanParameter("Headers?", "H?",
                "Treat the first row of the range as column headers. The header row is removed " +
                "from Values and returned separately as Names.",
                GH_ParamAccess.item, false);
            _inDirection = pManager.AddIntegerParameter("Direction", "D",
                "Tree orientation. 0 = By Row (one branch per spreadsheet row), " +
                "1 = By Column (one branch per spreadsheet column).",
                GH_ParamAccess.item, 0);
            _inFormulas = pManager.AddBooleanParameter("Formulas?", "F?",
                "If true, cells containing a formula yield the formula text (e.g. '=B2*2') " +
                "instead of the cached result.",
                GH_ParamAccess.item, false);
            _inCache = pManager.AddBooleanParameter("Cache?", "C?",
                "Cache the parsed range so unrelated Grasshopper solves do not re-read the file. " +
                "The cache invalidates automatically when the workbook changes on disk.",
                GH_ParamAccess.item, true);

            if (pManager[_inDirection] is Param_Integer direction)
            {
                direction.AddNamedValue("By Row", 0);
                direction.AddNamedValue("By Column", 1);
            }

            pManager[_inSheet].Optional = true;
            pManager[_inRange].Optional = true;
            pManager[_inExpand].Optional = true;
            pManager[_inHeaders].Optional = true;
            pManager[_inDirection].Optional = true;
            pManager[_inFormulas].Optional = true;
            pManager[_inCache].Optional = true;
        }

        private int _outValues;
        private int _outNames;
        private int _outAddress;
        private int _outRows;
        private int _outColumns;

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            _outValues = pManager.AddGenericParameter("Values", "V",
                "Cell values as a data tree. Numbers, dates, booleans and text keep their type; " +
                "empty cells are null.",
                GH_ParamAccess.tree);
            _outNames = pManager.AddTextParameter("Names", "N",
                "Column headers taken from the first row of the range, in column order. " +
                "Empty unless Headers? is true.",
                GH_ParamAccess.list);
            _outAddress = pManager.AddTextParameter("Address", "A",
                "The address actually read, e.g. 'Data!B2:D11'.",
                GH_ParamAccess.item);
            _outRows = pManager.AddIntegerParameter("Rows", "R",
                "Number of data rows returned (excluding the header row when Headers? is true).",
                GH_ParamAccess.item);
            _outColumns = pManager.AddIntegerParameter("Columns", "C",
                "Number of spreadsheet columns read.",
                GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            var read = true;
            da.GetData(_inRead, ref read);
            if (!read)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Read is false; no action taken.");
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
            var expand = true;
            var headers = false;
            var direction = 0;
            var formulas = false;
            var cache = true;
            da.GetData(_inSheet, ref sheet);
            da.GetData(_inRange, ref range);
            da.GetData(_inExpand, ref expand);
            da.GetData(_inHeaders, ref headers);
            da.GetData(_inDirection, ref direction);
            da.GetData(_inFormulas, ref formulas);
            da.GetData(_inCache, ref cache);

            if (string.IsNullOrWhiteSpace(range)) range = "A1";

            try
            {
                if (!TryResolveAddress(path, range, sheet, out var address, out var addressError))
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, addressError);
                    return;
                }

                var result = ExcelReadCache.GetOrRead(path, address, expand, formulas, cache);

                var grid = result.Grid;
                var names = new List<string>();

                if (headers && grid.Length > 0)
                {
                    // The header row is always the first *spreadsheet* row of the
                    // range, regardless of tree orientation. Names[c] therefore
                    // lines up with column c in either direction.
                    var headerRow = grid[0];
                    for (var c = 0; c < headerRow.Length; c++)
                        names.Add(HeaderName(headerRow[c], c));

                    grid = grid.Skip(1).ToArray();
                }

                da.SetDataTree(_outValues, BuildTree(grid, direction == 1));
                da.SetDataList(_outNames, names);
                da.SetData(_outAddress, result.Address);
                da.SetData(_outRows, grid.Length);
                da.SetData(_outColumns, result.ColumnCount);

                if (result.RowCount == 0 || result.ColumnCount == 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                        "The requested range is empty.");
                }
                else if (headers && grid.Length == 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                        "The range contains only a header row; no data rows were returned.");
                }

                // Formulas that Excel has never evaluated have no cached result,
                // so they read as null. This is the single most confusing
                // behaviour of the component, so call it out explicitly.
                if (!formulas && result.UncachedFormulaCount > 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                        $"{result.UncachedFormulaCount} formula cell(s) in this range have no cached result " +
                        $"(first: {result.FirstUncachedFormula}) and are returned as null. " +
                        "Excel stores formula results only when it recalculates, so formulas written by " +
                        "Ducky stay uncached until the workbook is opened in Excel. " +
                        "Set Formulas? to true to read the formula text instead.");
                }
            }
            catch (Exception ex)
            {
                ReportError("Excel read failed", ex);
            }
        }

        /// <summary>
        /// Converts a header cell into a usable column name, falling back to a
        /// positional name for blank headers so Names stays aligned with columns.
        /// </summary>
        private static string HeaderName(object value, int columnIndex)
        {
            if (value is null) return "col_" + columnIndex.ToString(CultureInfo.InvariantCulture);

            var text = value as string ?? Convert.ToString(value, CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(text)
                ? "col_" + columnIndex.ToString(CultureInfo.InvariantCulture)
                : text.Trim();
        }

        /// <summary>
        /// Parses the range and, when it names a defined range or Table,
        /// resolves that name against the workbook.
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

            // An explicit Sheet input fills in the gap when the range is unqualified.
            if (address.Sheet is null && !string.IsNullOrWhiteSpace(sheet))
                address = address.WithSheet(sheet);

            return true;
        }

        /// <summary>
        /// Converts the row-major grid into a data tree. By Row yields one
        /// branch per spreadsheet row; By Column transposes.
        /// </summary>
        private static GH_Structure<IGH_Goo> BuildTree(object[][] grid, bool byColumn)
        {
            var tree = new GH_Structure<IGH_Goo>();
            if (grid is null || grid.Length == 0) return tree;

            var rows = grid.Length;
            var columns = grid[0].Length;

            if (!byColumn)
            {
                for (var r = 0; r < rows; r++)
                {
                    var path = new GH_Path(r);
                    var branch = new List<IGH_Goo>(columns);
                    for (var c = 0; c < columns; c++)
                        branch.Add(TypeMapping.ToGoo(grid[r][c]));
                    tree.AppendRange(branch, path);
                }
                return tree;
            }

            for (var c = 0; c < columns; c++)
            {
                var path = new GH_Path(c);
                var branch = new List<IGH_Goo>(rows);
                for (var r = 0; r < rows; r++)
                    branch.Add(TypeMapping.ToGoo(grid[r][c]));
                tree.AppendRange(branch, path);
            }
            return tree;
        }
    }
}



