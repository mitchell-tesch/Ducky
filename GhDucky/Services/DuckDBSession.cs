using System;
using System.Data.Common;
using System.Threading;
using DuckDB.NET.Data;

namespace GhDucky.Services
{
    public sealed class DuckDBSession : IDisposable
    {
        private int _disposed;
        private readonly object _connectionLock = new();

        internal DuckDBSession(string id, string source, string displayName, bool isInMemory)
        {
            Id = id;
            Source = source;
            DisplayName = string.IsNullOrWhiteSpace(displayName) ? id : displayName;
            IsInMemory = isInMemory;

            var builder = new DbConnectionStringBuilder
            {
                ["DataSource"] = isInMemory ? ":memory:" : source
            };

            Connection = new DuckDBConnection(builder.ConnectionString);
            
            // established practice for .NET UI: wrap unmanaged blocking calls in a task with timeout
            // to prevent the host application (Rhino) from hanging if the file is locked or on a 
            // dropped network share.
            var openTask = System.Threading.Tasks.Task.Run(() => Connection.Open());
            if (!openTask.Wait(TimeSpan.FromSeconds(15)))
            {
                throw new TimeoutException($"Connection to {DisplayName} timed out after 15 seconds. The file may be locked by another process or the network location is inaccessible.");
            }

            CreatedAt = DateTime.UtcNow;
        }

        public string Id { get; }
        public string DisplayName { get; }
        public string Source { get; }
        public bool IsInMemory { get; }
        internal DuckDBConnection Connection { get; }
        public DateTime CreatedAt { get; }

        public bool IsDisposed => Volatile.Read(ref _disposed) != 0;

        public bool IsOpen => !IsDisposed && Connection.State == System.Data.ConnectionState.Open;

        /// <summary>
        /// Executes an action against the connection under a lock so that
        /// concurrent Grasshopper component solves do not interleave commands.
        /// </summary>
        public void Execute(Action<DuckDBConnection> action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            ObjectDisposedException.ThrowIf(IsDisposed, this);
            lock (_connectionLock)
            {
                ObjectDisposedException.ThrowIf(IsDisposed, this);
                action(Connection);
            }
        }

        /// <summary>
        /// Executes a function against the connection under a lock and returns
        /// its result.
        /// </summary>
        public T Execute<T>(Func<DuckDBConnection, T> func)
        {
            if (func == null) throw new ArgumentNullException(nameof(func));
            ObjectDisposedException.ThrowIf(IsDisposed, this);
            lock (_connectionLock)
            {
                ObjectDisposedException.ThrowIf(IsDisposed, this);
                return func(Connection);
            }
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0)
                return;

            try
            {
                if (Connection.State != System.Data.ConnectionState.Closed)
                    Connection.Close();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError($"DuckDBSession: Failed to close connection for {DisplayName}. {ex}");
                // best-effort close
            }

            Connection.Dispose();
        }
    }
}
