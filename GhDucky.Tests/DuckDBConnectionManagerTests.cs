using System;
using System.IO;
using GhDucky.Services;
using Xunit;

namespace GhDucky.Tests;

public class DuckDBConnectionManagerTests
{
    [Fact]
    public void Open_MultipleTimes_ReturnsSameSession()
    {
        // Ensure clean state
        DuckDBConnectionManager.CloseAll();

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".duckdb");
        try
        {
            var s1 = DuckDBConnectionManager.Open(path, "test");
            var s2 = DuckDBConnectionManager.Open(path, "test");

            Assert.Same(s1, s2);
            Assert.False(s1.IsDisposed);
        }
        finally
        {
            DuckDBConnectionManager.CloseAll();
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Close_DecrementsRef_OnlyDisposesOnZero()
    {
        // Ensure clean state
        DuckDBConnectionManager.CloseAll();

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".duckdb");
        try
        {
            var s1 = DuckDBConnectionManager.Open(path, "test");
            var s2 = DuckDBConnectionManager.Open(path, "test");

            Assert.Same(s1, s2);

            // First close should NOT dispose the session because s2 is still holding it.
            DuckDBConnectionManager.Close(s1.Id);
            Assert.False(s1.IsDisposed, "Session should still be open after first Close because of reference counting.");
            
            // Checking TryGet should still return the session
            Assert.True(DuckDBConnectionManager.TryGet(s1.Id, out _));

            // Second close should finally dispose it.
            DuckDBConnectionManager.Close(s1.Id);
            Assert.True(s1.IsDisposed, "Session should be disposed after final Close.");
            Assert.False(DuckDBConnectionManager.TryGet(s1.Id, out _));
        }
        finally
        {
            DuckDBConnectionManager.CloseAll();
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void AnonymousSessions_AlwaysUnique()
    {
        DuckDBConnectionManager.CloseAll();

        var s1 = DuckDBConnectionManager.Open(null, null);
        var s2 = DuckDBConnectionManager.Open(null, null);

        Assert.NotEqual(s1.Id, s2.Id);
        Assert.NotSame(s1, s2);
        
        DuckDBConnectionManager.Close(s1.Id);
        DuckDBConnectionManager.Close(s2.Id);
    }

    [Fact]
    public void CloseAll_OverridesRefCounts()
    {
        DuckDBConnectionManager.CloseAll();
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".duckdb");
        try
        {
            var s1 = DuckDBConnectionManager.Open(path, "test");
            var s2 = DuckDBConnectionManager.Open(path, "test"); // Ref count = 2

            DuckDBConnectionManager.CloseAll();

            Assert.True(s1.IsDisposed);
            Assert.False(DuckDBConnectionManager.TryGet(s1.Id, out _));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Concurrent_OpenClose_StaysConsistent()
    {
        DuckDBConnectionManager.CloseAll();
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".duckdb");
        const int iterations = 1000;

        try
        {
            System.Threading.Tasks.Parallel.For(0, iterations, _ =>
            {
                var session = DuckDBConnectionManager.Open(path, "concurrent");
                DuckDBConnectionManager.Close(session.Id);
            });

            // After equal opens and closes, it should be disposed and gone.
            var active = DuckDBConnectionManager.ActiveSessions();
            Assert.Empty(active);
        }
        finally
        {
            DuckDBConnectionManager.CloseAll();
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
