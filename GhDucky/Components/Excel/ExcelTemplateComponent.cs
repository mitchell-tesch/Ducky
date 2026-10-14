using System;
using System.Collections.Generic;
using System.IO;
using GhDucky.Services;
using GhDucky.Utils;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
namespace GhDucky.Components.Excel
{
    /// <summary>
    /// Renders a styled Excel template by substituting placeholder values.
    /// </summary>
    public class ExcelTemplateComponent : DuckyComponentBase
    {
        public ExcelTemplateComponent()
            : base(
                "Ducky Excel Template",
                "DuckyXlTmpl",
                "Fills a styled Excel template (.xlsx/.xlsm) with data and saves it to a new file. " +
                "The template's formatting, formulas, charts and layout are preserved - design the " +
                "report in Excel, supply only the values from Grasshopper. " +
                "Use '{{key}}' for single values and '{{table.column}}' for repeating rows.",
                "Ducky",
                "5 | Excel")
        {
        }
        public override Guid ComponentGuid => new Guid("6e21d8af-94c7-4b03-a5f2-3d7be0c15a47");
        public override GH_Exposure Exposure => GH_Exposure.tertiary;
        protected override System.Drawing.Bitmap Icon => IconFactory.Build("📝", IconFactory.Excel);
        private int _inFill;
        private int _inTemplate;
        private int _inPath;
        private int _inKeys;
        private int _inValues;
        private int _inOverwrite;
        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            _inFill = pManager.AddBooleanParameter("Fill?", "F?",
                "Set to true to render the template.",
                GH_ParamAccess.item, false);
            _inTemplate = pManager.AddTextParameter("Template", "T",
                "Absolute path to the template workbook containing '{{...}}' placeholders.",
                GH_ParamAccess.item);
            _inPath = pManager.AddTextParameter("Path", "P",
                "Absolute path for the rendered output workbook. Must differ from the template.",
                GH_ParamAccess.item);
            _inKeys = pManager.AddTextParameter("Keys", "K",
                "Placeholder names, parallel to the branches of Values. " +
                "A plain name like 'title' fills '{{title}}' with a single value. " +
                "A dotted name like 'rows.qty' contributes the column 'qty' to the repeating " +
                "collection '{{rows.qty}}'; several dotted keys sharing a prefix form one table.",
                GH_ParamAccess.list);
            _inValues = pManager.AddGenericParameter("Values", "V",
                "Data tree with one branch per entry in Keys. For a plain key only the first " +
                "item is used; for a dotted key the whole branch becomes that column's rows.",
                GH_ParamAccess.tree);
            _inOverwrite = pManager.AddBooleanParameter("Overwrite?", "O?",
                "Overwrite the output file if it already exists.",
                GH_ParamAccess.item, true);
            pManager[_inOverwrite].Optional = true;
        }
        private int _outPath;
        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            _outPath = pManager.AddTextParameter("Path", "P",
                "Path to the rendered workbook.",
                GH_ParamAccess.item);
        }
        protected override void SolveInstance(IGH_DataAccess da)
        {
            var fill = false;
            da.GetData(_inFill, ref fill);
            if (!fill)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Fill is false; no action taken.");
                return;
            }
            var template = string.Empty;
            if (!da.GetData(_inTemplate, ref template) || string.IsNullOrWhiteSpace(template))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Template path required.");
                return;
            }
            if (!File.Exists(template))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Template not found: " + template);
                return;
            }
            var path = string.Empty;
            if (!da.GetData(_inPath, ref path) || string.IsNullOrWhiteSpace(path))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Output path required.");
                return;
            }
            if (!ExcelFileGuard.IsSupportedExtension(template) || !ExcelFileGuard.IsSupportedExtension(path))
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
            var keys = new List<string>();
            da.GetDataList(_inKeys, keys);
            if (keys.Count == 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "At least one key is required.");
                return;
            }
            if (!da.GetDataTree(_inValues, out GH_Structure<IGH_Goo> tree) || tree is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "No values supplied.");
                return;
            }
            var overwrite = true;
            da.GetData(_inOverwrite, ref overwrite);
            var branches = tree.Branches;
            if (branches.Count != keys.Count)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning,
                    $"{keys.Count} key(s) but {branches.Count} branch(es) in Values; " +
                    "they are matched by index, so any excess is ignored.");
            }
            try
            {
                var columns = new List<IReadOnlyList<object>>(keys.Count);
                for (var i = 0; i < keys.Count; i++)
                {
                    if (i >= branches.Count)
                    {
                        columns.Add(Array.Empty<object>());
                        continue;
                    }
                    var branch = branches[i];
                    var column = new List<object>(branch?.Count ?? 0);
                    if (branch != null)
                    {
                        foreach (var goo in branch)
                            column.Add(TypeMapping.Unwrap(goo));
                    }
                    columns.Add(column);
                }
                var values = ExcelTemplate.BuildValues(keys, columns);
                ExcelTemplate.Fill(template, path, values, overwrite);
                da.SetData(_outPath, Path.GetFullPath(path));
            }
            catch (Exception ex)
            {
                ReportError("Excel template fill failed", ex);
            }
        }
    }
}