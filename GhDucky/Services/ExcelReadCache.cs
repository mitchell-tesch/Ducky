using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using GhDucky.Utils;

namespace GhDucky.Services
{
    /// <summary>
    /// Caches parsed range reads so that a Grasshopper solve triggered by an
    /// unrelated upstream change (dragging a slider, moving geometry) does not
    /// re-parse the workbook.
    /// <para>
    /// Entries are keyed by the file's identity <i>and</i> its last-write stamp
    /// and length, so editing the workbook on disk invalidates the cache with
    /// no explicit signal. Writes through <see cref="ExcelWriter"/> invalidate
    /// eagerly as well, because a write within the same filesystem timestamp
    /// granularity could otherwise go unnoticed.
    /// </para>
    /// </summary>
    internal static class ExcelReadCache
    {
        /// <summary>
        /// Upper bound on cached reads. Excel ranges can be large, so this is
        /// kept deliberately small; eviction is oldest-first.
        /// </summary>
        private const int MaxEntries = 32;

        private sealed class Entry
        {
            public ExcelReadResult Result;
            public long Sequence;
        }

        private static readonly ConcurrentDictionary<string, Entry> Cache = new(StringComparer.Ordinal);
        private static long _sequence;

        /// <summary>
        /// Returns a cached read when available, otherwise performs the read and
        /// caches it.
        /// </summary>
        public static ExcelReadResult GetOrRead(
            string path, ExcelAddress address, bool expand, bool formulaText, bool useCache)
        {
            if (!useCache)
                return ExcelReader.Read(path, address, expand, formulaText);

            var key = BuildKey(path, address, expand, formulaText);
            if (key is null)
                return ExcelReader.Read(path, address, expand, formulaText);

            if (Cache.TryGetValue(key, out var hit))
            {
                hit.Sequence = System.Threading.Interlocked.Increment(ref _sequence);
                return hit.Result;
            }

            var result = ExcelReader.Read(path, address, expand, formulaText);

            Cache[key] = new Entry
            {
                Result = result,
                Sequence = System.Threading.Interlocked.Increment(ref _sequence),
            };

            Evict();
            return result;
        }

        /// <summary>Drops every cached read for the given workbook.</summary>
        public static void Invalidate(string path)
        {
            string full;
            try { full = Path.GetFullPath(path); }
            catch { return; }

            var prefix = full.ToUpperInvariant() + "|";
            foreach (var key in Cache.Keys.Where(k => k.StartsWith(prefix, StringComparison.Ordinal)).ToList())
                Cache.TryRemove(key, out _);
        }

        /// <summary>Drops the entire cache (used by tests and on plugin teardown).</summary>
        public static void Clear() => Cache.Clear();

        /// <summary>
        /// Builds a key incorporating the file's mutation stamp. Returns null if
        /// the file cannot be stat'ed, in which case caching is skipped.
        /// </summary>
        private static string BuildKey(string path, ExcelAddress address, bool expand, bool formulaText)
        {
            try
            {
                var info = new FileInfo(Path.GetFullPath(path));
                if (!info.Exists) return null;

                var extent = address.EndRow is null || address.EndColumn is null
                    ? "open"
                    : $"{address.EndRow}:{address.EndColumn}";

                return string.Join("|",
                    info.FullName.ToUpperInvariant(),
                    info.LastWriteTimeUtc.Ticks.ToString(CultureInfo.InvariantCulture),
                    info.Length.ToString(CultureInfo.InvariantCulture),
                    address.Sheet ?? string.Empty,
                    address.Name ?? string.Empty,
                    address.StartRow.ToString(CultureInfo.InvariantCulture),
                    address.StartColumn.ToString(CultureInfo.InvariantCulture),
                    extent,
                    expand ? "1" : "0",
                    formulaText ? "1" : "0");
            }
            catch
            {
                return null;
            }
        }

        private static void Evict()
        {
            if (Cache.Count <= MaxEntries) return;

            foreach (var key in Cache
                         .OrderBy(kv => kv.Value.Sequence)
                         .Take(Cache.Count - MaxEntries)
                         .Select(kv => kv.Key)
                         .ToList())
            {
                Cache.TryRemove(key, out _);
            }
        }
    }
}

