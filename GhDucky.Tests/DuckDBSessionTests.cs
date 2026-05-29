using System;
using System.IO;
using GhDucky.Services;
using Xunit;

namespace GhDucky.Tests;

public class DuckDBSessionTests
{
    [Fact]
    public void Constructor_EscapesDataSource_WhenPathContainsSpecialCharacters()
    {
        // Use a path that would break simple string concatenation
        var maliciousPath = Path.Combine(Path.GetTempPath(), "test;access_mode=read_write.duckdb");
        var id = Guid.NewGuid().ToString("N");

        try
        {
            using var session = new DuckDBSession(id, maliciousPath, "test", false);
            
            // The ConnectionString should have the path properly escaped/quoted.
            // DbConnectionStringBuilder usually wraps values with special characters in double quotes.
            var connectionString = session.Connection.ConnectionString;
            
            Assert.Contains($"\"{maliciousPath}\"", connectionString);
            
            // Ensure it didn't actually inject a separate access_mode key (though this is harder to assert 
            // without parsing the connection string back, which DbConnectionStringBuilder can do).
            var parser = new System.Data.Common.DbConnectionStringBuilder { ConnectionString = connectionString };
            Assert.Equal(maliciousPath, parser["DataSource"]);
            Assert.Single(parser.Keys);
        }
        finally
        {
            if (File.Exists(maliciousPath)) File.Delete(maliciousPath);
        }
    }

    [Fact]
    public void Constructor_InMemory_UsesMemoryDataSource()
    {
        var id = Guid.NewGuid().ToString("N");
        using var session = new DuckDBSession(id, ":memory:", "test", true);
        
        var parser = new System.Data.Common.DbConnectionStringBuilder { ConnectionString = session.Connection.ConnectionString };
        Assert.Equal(":memory:", parser["DataSource"]);
    }
}
