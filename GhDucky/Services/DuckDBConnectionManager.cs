using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;

namespace GhDucky.Services
{
    /// Process-wide registry of open DuckDB sessions, keyed by a stable ID
    /// so handles can flow through Grasshopper components without re-opening
    /// the database or losing in-memory state.
    public static class DuckDBConnectionManager
    {
        private static readonly ConcurrentDictionary<string, DuckDBSession> Sessions =
            new(StringComparer.Ordinal);

        private static readonly ConcurrentDictionary<string, int> RefCounts =
            new(StringComparer.Ordinal);

        private static readonly object SourceLock = new();
        private static readonly Dictionary<string, string> SourceToId =
            new(StringComparer.OrdinalIgnoreCase);

        /// Opens (or returns) a session.
        /// <paramref name="source"/> empty/null/":memory:" -> in-memory database.
        /// File-backed sources are deduplicated by absolute path.
        /// In-memory sources are deduplicated by <paramref name="displayName"/> when supplied,
        /// otherwise a fresh session is created.
        public static DuckDBSession Open(string source, string displayName)
        {
            NativeLibraryResolver.Initialize();

            var isInMemory = string.IsNullOrWhiteSpace(source) ||
                             string.Equals(source, ":memory:", StringComparison.OrdinalIgnoreCase);

            string canonicalSource;
            string key;

            if (isInMemory)
            {
                canonicalSource = ":memory:";
                key = string.IsNullOrWhiteSpace(displayName)
                    ? null
                    : "mem::" + displayName;
            }
            else
            {
                canonicalSource = Path.GetFullPath(source);
                key = "file::" + canonicalSource;
            }

            if (key != null)
            {
                lock (SourceLock)
                {
                    if (SourceToId.TryGetValue(key, out var existingId) &&
                        Sessions.TryGetValue(existingId, out var existing) &&
                        existing.IsOpen)
                    {
                        RefCounts.AddOrUpdate(existingId, 1, (_, count) => count + 1);
                        return existing;
                    }

                    var id = Guid.NewGuid().ToString("N");
                    var name = string.IsNullOrWhiteSpace(displayName)
                        ? (isInMemory ? "memory" : Path.GetFileName(canonicalSource))
                        : displayName;

                    var session = new DuckDBSession(id, canonicalSource, name, isInMemory);
                    Sessions[id] = session;
                    SourceToId[key] = id;
                    RefCounts[id] = 1;
                    return session;
                }
            }

            // Anonymous in-memory: always fresh, not shared, but still registered
            // in Sessions so that GH_DuckDBConnection.Session (which uses TryGet)
            // can resolve it.  Callers must Close(id) when done.
            var anonId = Guid.NewGuid().ToString("N");
            var anonSession = new DuckDBSession(anonId, ":memory:", "memory", true);
            Sessions[anonId] = anonSession;
            RefCounts[anonId] = 1;
            return anonSession;
        }

        public static bool TryGet(string id, out DuckDBSession session)
        {
            if (string.IsNullOrEmpty(id))
            {
                session = null;
                return false;
            }

            if (Sessions.TryGetValue(id, out session) && session.IsOpen)
                return true;

            session = null;
            return false;
        }

        public static void Close(string id)
        {
            if (string.IsNullOrEmpty(id))
                return;

            lock (SourceLock)
            {
                if (!RefCounts.TryGetValue(id, out var count))
                    return;

                if (count > 1)
                {
                    RefCounts[id] = count - 1;
                    return;
                }

                // Final reference: clean up and dispose.
                if (!Sessions.TryRemove(id, out var session))
                    return;

                DisposeSessionInternal(id, session);

                string keyToRemove = null;
                foreach (var kv in SourceToId)
                {
                    if (kv.Value == id)
                    {
                        keyToRemove = kv.Key;
                        break;
                    }
                }
                if (keyToRemove != null)
                    SourceToId.Remove(keyToRemove);
            }
        }

        private static void DisposeSessionInternal(string id, DuckDBSession session)
        {
            RefCounts.TryRemove(id, out _);
            DuckDbExtensionTracker.ClearAllForSession(id);
            session.Dispose();
        }

        public static IReadOnlyCollection<DuckDBSession> ActiveSessions()
        {
            return Sessions.Values is ICollection<DuckDBSession> col
                ? (IReadOnlyCollection<DuckDBSession>)new List<DuckDBSession>(col)
                : new List<DuckDBSession>();
        }

        /// <summary>
        /// Closes and disposes every active session. Useful for clean teardown
        /// when Grasshopper or the host process is shutting down so that file
        /// locks are released and native handles are freed.
        /// </summary>
        public static void CloseAll()
        {
            lock (SourceLock)
            {
                var ids = new List<string>(Sessions.Keys);
                foreach (var id in ids)
                {
                    try
                    {
                        if (!Sessions.TryRemove(id, out var session))
                            continue;

                        DisposeSessionInternal(id, session);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Trace.TraceError($"DuckDBConnectionManager: Failed to close session {id} during teardown. {ex}");
                    }
                }
                SourceToId.Clear();
            }
        }
    }
}
