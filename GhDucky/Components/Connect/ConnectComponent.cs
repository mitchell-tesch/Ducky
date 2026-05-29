using System;
using GhDucky.Parameters;
using GhDucky.Services;
using GhDucky.Utils;
using Grasshopper.Kernel;

namespace GhDucky.Components.Connect
{
    public class ConnectComponent : DuckyComponentBase
    {
        public ConnectComponent()
            : base(
                "Ducky Connect",
                "DuckyOn",
                "Opens or creates a Database session.\n" +
                "Leave the source empty (or supply ':memory:') for an in-memory database.\n" +
                "Supplying a file path opens (and creates if missing) a persistent database file.",
                "Ducky",
                "1 | Connect")
        {
        }

        public override Guid ComponentGuid => new Guid("d6c8b6ae-3a4f-4f0f-aa14-3a8e4c4dbb1a");

        public override GH_Exposure Exposure => GH_Exposure.primary;

        protected override System.Drawing.Bitmap Icon => IconFactory.Build("🦆", IconFactory.Connect);
        
        private int _inOpen;
        private int _inSource;
        private int _inName;

        // Per-component cache so anonymous in-memory connections survive
        // recomputes (the manager cannot dedupe them without a name, so without
        // this cache every solve would create — and orphan — a fresh in-memory
        // database).
        private string _cachedSessionId;
        private string _cachedSource;
        private string _cachedName;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            _inOpen = pManager.AddBooleanParameter(
                "Connect?", "C?",
                "Set to true to open the connection.",
                GH_ParamAccess.item, true);
            _inSource = pManager.AddTextParameter(
                "Source", "S",
                "Database source.\nEmpty or ':memory:' for an in-memory database; otherwise a path to a .duckdb file.",
                GH_ParamAccess.item, string.Empty);
            _inName = pManager.AddTextParameter(
                "Name", "N",
                "Optional display name.\nFor in-memory databases, supplying a name lets multiple components share the same in-memory store.",
                GH_ParamAccess.item, string.Empty);

            pManager[_inSource].Optional = true;
            pManager[_inName].Optional = true;
        }

        private int _outDatabase;
        private int _outInfo;
        
        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            _outDatabase = pManager.AddParameter(new ParamDuckyDbConnection(), "Database", "DB",
                "Database connection.", 
                GH_ParamAccess.item);
            _outInfo = pManager.AddTextParameter("Info", "I",
                "Connection status information.", 
                GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            var open = true;
            da.GetData(_inOpen, ref open);
            if (!open)
            {
                ReleaseCachedSession();
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Connection is closed, set Open to true to open.");
                return;
            }
            
            var source = string.Empty;
            var name = string.Empty;
            da.GetData(_inSource, ref source);
            da.GetData(_inName, ref name);
            try
            {
                var session = ReuseOrOpen(source, name);
                da.SetData(_outDatabase, new GH_DuckDBConnection(session));
                da.SetData(_outInfo, BuildInfo(session));
            }
            catch (Exception ex)
            {
                ReportError("Failed to open connection", ex);
            }
        }

        /// <summary>
        /// Returns the cached session for this component when the (source, name)
        /// inputs are unchanged and the session is still open; otherwise opens a
        /// new session, closing the reference to any previously cached session.
        /// </summary>
        private DuckDBSession ReuseOrOpen(string source, string name)
        {
            var inputsUnchanged =
                string.Equals(_cachedSource ?? string.Empty, source ?? string.Empty, StringComparison.Ordinal) &&
                string.Equals(_cachedName ?? string.Empty, name ?? string.Empty, StringComparison.Ordinal);

            if (inputsUnchanged &&
                !string.IsNullOrEmpty(_cachedSessionId) &&
                DuckDBConnectionManager.TryGet(_cachedSessionId, out var cached))
            {
                return cached;
            }

            // Inputs changed (or no live cache). Release the old session reference first.
            ReleaseCachedSession();

            var session = DuckDBConnectionManager.Open(source, name);
            _cachedSessionId = session.Id;
            _cachedSource = source ?? string.Empty;
            _cachedName = name ?? string.Empty;
            return session;
        }

        private void ReleaseCachedSession()
        {
            if (string.IsNullOrEmpty(_cachedSessionId))
                return;

            // The manager handles reference counting: the session is only
            // physically closed and disposed when the last component using
            // it calls Close().
            DuckDBConnectionManager.Close(_cachedSessionId);
            _cachedSessionId = null;
        }

        public override void RemovedFromDocument(GH_Document document)
        {
            try { ReleaseCachedSession(); }
            catch { /* best-effort cleanup */ }
            base.RemovedFromDocument(document);
        }
        
        private static string BuildInfo(DuckDBSession session)
        {
            var kind = session.IsInMemory ? "in-memory" : "file";
            return $"DuckyDB session [{session.DisplayName}] ({kind}) id={session.Id}\nSource: {session.Source}";
        }
    }
}
