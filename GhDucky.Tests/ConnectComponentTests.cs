using System;
using System.Reflection;
using GhDucky.Components.Connect;
using GhDucky.Services;
using Xunit;

namespace GhDucky.Tests;

public class ConnectComponentTests
{
    [Fact]
    public void ConnectComponent_Integration_DecrementsOnRemoval()
    {
        DuckDBConnectionManager.CloseAll();

        // 1. Setup a session in the manager.
        var session = DuckDBConnectionManager.Open(null, "integration-test");
        var sessionId = session.Id;

        // 2. Create the component.
        var component = new ConnectComponent();

        // 3. Use reflection to simulate that the component has "opened" this session.
        // We set the private field _cachedSessionId.
        var field = typeof(ConnectComponent).GetField("_cachedSessionId", 
            BindingFlags.NonPublic | BindingFlags.Instance);
        
        if (field == null)
        {
            // Fallback for base class if it moved
            field = typeof(ConnectComponent).BaseType.GetField("_cachedSessionId", 
                BindingFlags.NonPublic | BindingFlags.Instance);
        }
        
        Assert.NotNull(field);
        field.SetValue(component, sessionId);

        // 4. Act: Remove from document (simulates deletion in GH).
        component.RemovedFromDocument(null);

        // 5. Assert: The session should be disposed because the ref count hit 0.
        Assert.True(session.IsDisposed, "Session should be disposed after component removal.");
        Assert.False(DuckDBConnectionManager.TryGet(sessionId, out _), "Session should be removed from manager.");
    }
}
