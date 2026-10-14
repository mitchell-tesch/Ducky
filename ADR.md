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

---

## ADR 2: Two Libraries for Direct Excel Read/Write

### Status
Accepted (2026-09-23)

### Context
The `5 | Excel` components provide cell- and range-level access to `.xlsx`/`.xlsm` workbooks, independent of DuckDB. The requirements were: fast, robust, no COM/interop, single cell or range, auto-expanding reads and writes, selectable orientation, in-place writes that preserve the rest of the workbook, formula support, and named-range addressing. Ducky is MIT-licensed, so any dependency must be MIT- or Apache-2.0-compatible.

Candidate libraries were assessed against those requirements:

- **EPPlus** — excluded on licence grounds (PolyForm Noncommercial from v5; LGPL at v4).
- **ClosedXML** (MIT) — meets every requirement with the least code, but 0.104+ depends on `SixLabors.Fonts`, which is under the Six Labors Split License (commercial use above a revenue threshold requires a paid licence). Passing that obligation on to users of an MIT plugin was judged unacceptable. It also loads the whole workbook into memory.
- **NPOI** (Apache-2.0) — capable, but heavier, and legacy `.xls` support was not required.
- **MiniExcel 2.0** (Apache-2.0) — evaluated empirically rather than from documentation, by probing the real API of `2.0.0-preview.4` against a hand-crafted workbook.
- **DocumentFormat.OpenXml** (MIT) — full control, no transitive licence ambiguity, but the lowest-level option.

The MiniExcel v2 probe established, concretely:

| Capability | Result |
|---|---|
| A1 range read (`QueryRange`) | Works |
| Open-ended range (null end row/column) | Works — expands to the used range |
| Native cell typing (double / bool / DateTime) | Works |
| Formula cells → cached value | Works |
| `GetSheetDimensions` / `GetSheetInformations` (used range, hidden state) | Works |
| Excel Tables by name (`QueryTable`) | Works |
| Placeholder templating (`FillTemplate`) | Works |
| **Cell-level in-place write** | **Not supported** |
| **Named ranges (`definedNames`)** | **Not supported** |
| **Formula text on read / formula writing** | **Not supported** |

The decisive test: writing a single value into an existing sheet via `InsertSheet(..., overwriteSheet: true)` reduced a sheet spanning `A1:C6`, including a formula, to a single populated cell. MiniExcel's only write verbs (`Export`, `InsertSheet`, `CopyAndAddSheet`, `AlterSheet`, `FillTemplate`) operate on whole files or whole sheets. Adding a *new* sheet is non-destructive; modifying an existing one is not.

### Decision
Use **MiniExcel.OpenXml for reads and templating** and **DocumentFormat.OpenXml for writes**.

MiniExcel covers the read path almost completely, does so with streaming performance and low memory, and supplies a mature template engine. The OpenXML SDK supplies the one thing MiniExcel structurally cannot: anchored, style-preserving, in-place cell edits.

Two gaps are filled by reading the xlsx package directly, in `ExcelWorkbookMetadata`, rather than by adding a third dependency:
- **Defined names** are parsed from `xl/workbook.xml`.
- **Excel Tables** are enumerated from the worksheet relationship parts.

An xlsx is a zip of XML; these particular parts are small and schema-stable, so ~200 lines of `ZipArchive` + `XDocument` is a better trade than another dependency.

### Consequences
- **Pros:**
  - Every stated requirement is met, including in-place writes, formula read/write and named ranges.
  - All licences (Apache-2.0, MIT) are compatible with Ducky's MIT licence, with no split-licence obligations passed to users.
  - Reads stream and stay off the large-object heap; writes touch only target cells, so they are faster than a full file regenerate.
  - Templating came essentially free, which removed the need to design and expose a full cell-styling API.
- **Cons:**
  - Two independent xlsx implementations ship in the `.gha`, adding roughly 7.5 MB (the OpenXML SDK alone is ~6 MB).
  - MiniExcel 2.0 is a **preview**. The API may shift before release. It is confined to `ExcelReader`, `ExcelTemplate` and `ExcelWorkbookMetadata` so the blast radius of a breaking change is small.
  - The write path is more code than ClosedXML would have required; it is covered by an end-to-end round-trip test suite that asserts non-destructiveness directly.

### Design Notes & Tradeoffs

#### Formula handling happens before trimming
Formulas written by Ducky carry no cached `<v>`, so MiniExcel reads them as `null`. If empty row/column trimming ran first, a formula-only column would be trimmed away before the formula text could be overlaid. `ExcelReader.Read` therefore resolves formulas against the full grid, then trims.

#### One stream per MiniExcel call
MiniExcel 2.0 preview does not consistently honour its `leaveOpen` parameter — `GetSheetInformations` closes the supplied stream regardless. Rather than depend on which methods happen to respect it, every MiniExcel call opens its own `FileStream`. These are opened with `FileShare.ReadWrite` so a workbook open in Excel can still be read.

#### Writes are staged, not edited in place on disk
The OpenXML SDK mutates the zip container directly. A crash or power loss mid-save would leave a corrupt workbook. Writes are therefore staged into a sibling temporary file and committed with `File.Replace`, which is atomic and preserves the destination's attributes and ACLs, with a copy-based fallback for network shares that reject it.

#### Staging files must not drive format decisions
Because writes are staged to a sibling `.tmp` file, anything derived from the file extension must use the *destination* path. This was originally missed for the macro-enabled workbook check, which inspected the staging path, saw `.tmp`, and produced a plain xlsx content type inside a `.xlsm` file — a mismatch Excel warns about on open. `DocumentTypeFor` now takes the destination path explicitly and is documented as such.

#### Inline strings over the shared string table
String cells are written as `<inlineStr>` rather than appended to the workbook-wide shared string table. Grasshopper re-solves frequently, and repeatedly appending to the shared string table would grow the file without bound. Inline strings keep writes idempotent at the cost of a marginally larger cell payload.

One consequence is worth knowing: MiniExcel's **template engine reads placeholders from the shared string table**, so a template authored by Ducky's own writer is not recognised. This is harmless in practice — templates are authored in Excel, which always uses shared strings — but it does mean template test fixtures must be built with shared strings to exercise the real path.

#### Uncached formulas are not empty cells
Excel persists a formula's result only when it recalculates, so a formula written by Ducky has an `<f>` but no `<v>`, and MiniExcel surfaces only the cached value. Two consequences are handled explicitly:

1. **Trimming.** A null formula cell would otherwise be trimmed away as a blank row or column, silently shrinking the returned range. `ExcelReader` builds an occupancy mask marking formula cells as occupied, and trimming consults it. The mask is allocated only when such a cell is found, so the common path stays allocation-free.
2. **Reporting.** The reader counts these cells and records the first one; the Read and Import Excel Range components warn and name it. Silently returning nulls was judged the most confusing behaviour in the component set.

Evaluating formulas in-process was considered and rejected: a correct Excel formula engine is a very large undertaking, and a partial one would produce results that disagree with Excel — worse than an honest null plus a warning.

#### Templating instead of a styling API
Formatted output is served by rendering a user-authored template rather than by exposing number formats, fonts, fills and widths as Grasshopper inputs. Design stays in Excel where it is far easier to author and iterate, and Ducky supplies only values. Dotted placeholder keys (`rows.name`, `rows.qty`) group into a single repeating collection, which matches Grasshopper's natural one-branch-per-column shape. A direct cell-styling API remains possible later, but is no longer urgent.


