# Architecture Decision Records (ADR)

## ADR 1: Database Session Lifecycle Management via Reference Counting

### Status
Accepted (2026-05-28)

### Context
DuckDB uses file-level locks to manage access to database files. In a Grasshopper environment, multiple components across one or more documents may reference the same database file. 

Previously, `DuckDBConnectionManager` deduplicated sessions by file path but had no mechanism to know when a session was no longer needed. This led to:
1. **Connection Leaks**: Sessions remained open in a static cache for the duration of the Rhino process.
2. **File Lock Exhaustion**: Users could not delete or move `.duckdb` files on disk even after deleting the Grasshopper components that referenced them, as the native file handle was still held by the process.

We considered two main approaches:
- **Option A: Strict Document Scoping**: Binding the connection lifecycle to the `GH_Document`. This is difficult because components can be copied between documents, and cross-document sharing of the same file (to save memory) is a valid use case.
- **Option B: Reference Counting**: Allowing the manager to track how many active "tokens" (held by components) exist for a session and disposing of it only when the count hits zero.

### Decision
We chose **Option B: Reference Counting** in `DuckDBConnectionManager`.

When a component "opens" a session via the manager, the reference count is incremented. When the component is removed or its inputs change, it calls `Close(id)`, which decrements the count. The manager only disposes the native `DuckDBConnection` and removes it from the global cache when the count reaches zero.

### Consequences
- **Pros**: 
  - Efficiently shares native memory and file locks across the entire process.
  - Automatically releases file locks when the last component using them is removed.
  - Transparent to the user; no manual "Disconnect" is strictly required for cleanup.
- **Cons**: 
  - Requires disciplined calling of `Close(id)` by all consumer components.
  - If a component fails to call `Close()` (e.g., due to an unhandled crash in a different part of the solve), a leak could still occur until the next `CloseAll()` (triggered by process shutdown).

### Design Notes & Tradeoffs
#### Intentional O(N) Lookup in `Close(id)`
In `DuckDBConnectionManager.Close(id)`, removing a session from the `SourceToId` map (the deduplication cache) requires an O(N) search through the dictionary to find the source key associated with that session ID. 

We intentionally chose this over adding a `SourceKey` property to the `DuckDBSession` object (which would allow O(1) removal). Adding the key to the session object would unnecessarily couple the core session representation to the manager's internal caching strategy. Given that a typical Grasshopper user will likely have fewer than 100 active database connections simultaneously, the O(N) lookup is functionally instantaneous and the architectural decoupling is a more valuable long-term benefit.
