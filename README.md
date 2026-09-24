# 🦆 Ducky — DuckDB for Grasshopper

**Ducky** brings the power of [DuckDB](https://duckdb.org/) to [Grasshopper](https://www.grasshopper3d.com/). Open in-memory or file-backed databases, load CSV/JSON/Parquet/Excel files and Grasshopper data trees, query everything in SQL, export results, and round-trip Rhino geometry through the DuckDB spatial extension.

[![Rhino 8](https://img.shields.io/badge/Rhino-8-blue)](https://www.rhino3d.com/)
[![Grasshopper](https://img.shields.io/badge/Grasshopper-8.0-green)](https://www.grasshopper3d.com/)
[![.NET 8](https://img.shields.io/badge/.NET-8.0-purple)](https://dotnet.microsoft.com/)
[![DuckDB](https://img.shields.io/badge/DuckDB-1.5.3-yellow)](https://duckdb.org/)

---

## Features

- **In-memory & file-backed databases** — spin up a throwaway analytics session or persist data to a `.duckdb` file.
- **Import CSV, JSON, Parquet, Excel** — leverage DuckDB's blazing-fast native readers (`read_csv_auto`, `read_json_auto`, `read_parquet`) and the `spatial`/`excel` extensions for `.xlsx` files.
- **Import Grasshopper data trees** — push branch data directly into tables via the high-performance Appender API with automatic type inference.
- **SQL queries** — run any SQL against your data; results come back as Grasshopper data trees with typed columns.
- **Export** — write query results or tables to CSV, JSON, Parquet, or Excel with a single component.
- **Direct Excel read/write** — cell- and range-level access to `.xlsx`/`.xlsm` workbooks, independent of the database. No Excel install, no COM, no interop. Auto-expanding ranges, row/column orientation, and in-place writes that preserve the rest of your workbook.
- **Spatial / Geometry round-trip** — import Rhino points, curves, meshes, and breps as WKB-encoded GEOMETRY columns via the DuckDB `spatial` extension, then query them back with `ST_*` functions.

---

## Components

### 1 · Connect

| Component | Nickname | Description |
|---|---|---|
| **Ducky Connect** | `DuckyOn` | Opens or creates a database session. Leave the source empty for in-memory; supply a file path for a persistent database. |
| **Ducky Disconnect** | `DuckyOff` | Closes a session and releases file locks. In-memory databases are discarded. |
| **Ducky Enable Spatial** | `DuckySpatial` | Installs and loads the DuckDB `spatial` extension on the connection. |
| **Ducky Enable Excel** | `DuckyExcel` | Installs and loads the DuckDB `spatial` and `excel` extensions for `.xlsx` support in manual SQL queries. Import/Export Excel components auto-enable these. |

### 2 · Import

| Component | Nickname | Description                                                                                   |
|---|---|-----------------------------------------------------------------------------------------------|
| **Ducky Import Data Tree** | `DuckyTree` | Imports a Grasshopper data tree into a table. Each branch is one column; items are row values. |
| **Ducky Import CSV** | `DuckyCSV` | Imports a CSV file into a table using `read_csv_auto`.                                        |
| **Ducky Import JSON** | `DuckyJSON` | Imports a JSON file (array or newline-delimited) into a table using `read_json_auto`.         |
| **Ducky Import Parquet** | `DuckyPQ` | Imports one or more Parquet files (supports globs) into a table.                              |
| **Ducky Import Excel** | `DuckyXLSX` | Imports an Excel (.xlsx) file into a table via the `spatial`/`excel` extensions. Supports sheet selection and headers. |
| **Ducky Import Excel Range** | `DuckyXlRng` | Imports any Excel cell range, defined name or Table into a table, so it can be queried in SQL. Not limited to whole sheets, and needs no Excel install. |
| **Ducky Import Geometry** | `DuckyGeo` | Imports Rhino geometry to a table as a WKB-encoded GEOMETRY column.                           |

### 3 · Query

| Component | Nickname | Description |
|---|---|---|
| **Ducky Inspect** | `DuckyPeek` | Lists tables, columns, types, and row counts for a database connection. |
| **Ducky Query** | `DuckyQ` | Executes a SQL query and returns the result as a data tree (one branch per column). |
| **Ducky Query Geometry** | `DuckyGeoQ` | Executes a SQL query that includes a GEOMETRY column and reconstructs Rhino geometry from WKB. |
| **Ducky Query Builder** | `DuckyQB` | Builds a SQL SELECT query from simple inputs — no SQL knowledge required. |
| **Ducky Filter** | `DuckyFlt` | Creates a filter condition for the Query Builder's Filters input. |
| **Ducky Join Tables** | `DuckyJoin` | Combines rows from two tables based on a shared column (JOIN). |

### 4 · Export

| Component | Nickname | Description |
|---|---|---|
| **Ducky Export** | `DuckyEx` | Exports a table or query result to CSV, JSON, or Parquet via DuckDB's `COPY TO`. |
| **Ducky Export Excel** | `DuckyExXL` | Exports a table or query result to an Excel (.xlsx) file via the `spatial`/`excel` extensions. |

### 5 · Excel

Direct, database-free access to workbooks. These components do **not** take a `Database` input and do **not** use Excel or COM — they read and write the OpenXML package itself, so they are fast and work headlessly.

| Component | Nickname | Description |
|---|---|---|
| **Ducky Excel Read** | `DuckyXlRead` | Reads a cell or range into a data tree. Supports A1 notation, sheet-qualified addresses, defined names and Excel Tables, auto-expansion from an anchor cell, optional header row, and row/column orientation. |
| **Ducky Excel Write** | `DuckyXlWrite` | Writes a data tree into a cell or range. In Place mode edits only the target cells, preserving all other cells, their formatting, the sheet's formulas and every other sheet. |
| **Ducky Excel Clear** | `DuckyXlClear` | Clears a cell, a range, or an entire sheet — values only, or values and formatting. |
| **Ducky Excel Template** | `DuckyXlTmpl` | Fills a styled Excel template with data and saves it to a new file, preserving the template's formatting, formulas and charts. |
| **Ducky Excel Info** | `DuckyXlInfo` | Inspects a workbook: sheet names, used ranges, dimensions, hidden state, defined names, Excel Tables, and file size / modified / locked state. |

#### Addressing

The **Range** input accepts any of:

| Form | Meaning |
|---|---|
| `B2` | Single anchor cell |
| `B2:D10` | Explicit range |
| `Data!B2:D10` | Sheet-qualified (overrides the **Sheet** input) |
| `'My Sheet'!B2` | Quoted sheet name |
| `$B$2:$D$10` | Absolute markers are accepted and ignored |
| `MyRange` | A defined name or Excel Table name, resolved against the workbook |

#### Reading

- **Expand?** — when only an anchor cell is given, the range grows to the enclosing **contiguous block**: a fully blank row bounds it vertically, a fully blank column bounds it horizontally (Excel's "current region"). With an explicit range, this is ignored.
- **Headers?** — treat the first row of the range as column names. The header row is removed from `Values` and returned separately as `Names`, in column order. Blank headers become `col_0`, `col_1`, …
- **Direction** — `By Row` gives one branch per spreadsheet row; `By Column` gives one branch per spreadsheet column. Branch paths are `{0}`, `{1}`, … in range order. `Names` is always in column order, so it lines up either way.
- **Types are preserved** — numbers arrive as numbers, dates as `GH_Time`, booleans as booleans, empty cells as `null`. Formula cells yield their cached result.
- **Formulas?** — set true to get the formula text (`=B2*2`) instead of the cached result.
- **Cache?** — the parsed range is cached between solves and invalidated automatically when the workbook changes on disk, so unrelated solves (slider drags, geometry edits) don't re-read the file.

> **⚠️ Formulas and cached results.** Excel stores a formula's *result* only when it recalculates. A formula written by Ducky therefore has no cached result and reads back as `null` until the workbook is opened in Excel. The Read component detects this and warns, naming the first offending cell. Set `Formulas?` to true to read the formula text instead. Such cells are never trimmed away as "empty".

#### Writing

| Mode | Behaviour |
|---|---|
| **In Place** (default) | Edits only the target cells. All other cells, their formatting, the sheet's formulas and other sheets are untouched. |
| **Overwrite Sheet** | Clears the target sheet first, then writes. Other sheets survive. |
| **New Workbook** | Replaces the file with a fresh single-sheet workbook. |

- **Auto-expanding write** — give just an anchor (`B2`) and the data tree's shape determines the extent. Give an explicit range and oversized data is clipped, with a warning naming the address actually written.
- **Nulls clear the cell** — the value is removed but the cell's formatting is kept.
- **Existing formatting is preserved** — number formats, fills, fonts and borders on overwritten cells survive. Dates get a date number format only when the cell doesn't already have one.
- **Formulas?** — when true (default), text beginning with `=` is written as a real formula. Excel is asked to fully recalculate on load, so formulas depending on written values refresh.
- **Missing files and sheets are created** automatically.
- **Crash-safe** — writes are staged to a sibling temp file and atomically swapped in, so an interrupted write can't corrupt the workbook.
- **Locked workbooks** are detected up front and reported as a clear "close the workbook" message rather than an exception.

#### Clearing

| Mode | Behaviour |
|---|---|
| **Values** | Clears cell contents and formulas, keeping the cells' formatting. |
| **Values and Formatting** | Removes the cells outright, discarding their formatting too. |
| **Entire Sheet** | Clears every cell on the sheet. The **Range** input is ignored; other sheets are untouched. |

Clearing never auto-expands — a single anchor cell clears exactly that cell.

#### Templates

Author the report in Excel, complete with fonts, colours, charts and formulas, and mark the data slots with placeholders:

| Placeholder | Supplied by |
|---|---|
| `{{title}}` | Key `title`; the first item of its branch |
| `{{rows.name}}`, `{{rows.qty}}` | Keys `rows.name` and `rows.qty`; the template row repeats once per item |

Dotted keys sharing a prefix form a single repeating table, so one branch per column is exactly what Grasshopper produces naturally. The output file must differ from the template, otherwise the placeholders would be consumed and the template destroyed — this is checked and refused.

> **Note:** `.xlsx` and `.xlsm` are supported; new `.xlsm` files are correctly created as macro-enabled workbooks. Legacy `.xls` is not supported.

---

## Getting Started

### Prerequisites

- [Rhino 8](https://www.rhino3d.com/) (Windows)
- .NET 8.0 SDK (for building from source)

### Install via Yak (Package Manager)

> search for **Ducky** in Rhino's package manager.

### Build from Source

```bash
git clone https://github.com/mitchell-tesch/Ducky.git
cd GhDucky
dotnet build
```

The build produces `Ducky.gha` in `GhDucky/bin/Debug/net8.0/`. Copy the entire output folder to your Grasshopper libraries directory:

```
%APPDATA%\Grasshopper\Libraries\Ducky\
```

> **Tip:** The build also auto-generates a Yak package if `Yak.exe` is found in your Rhino 8 installation.

### First Use

1. Drop a **Ducky Connect** component onto the canvas.
2. Set *Connect?* to `True` — this opens an in-memory database.
3. Wire the **Database** output into an import component (e.g. **Ducky Import CSV**).
4. Point the import at a file, set *Import?* to `True`.
5. Wire the **Database** output into a **Ducky Query** component, write your SQL, and set *Run?* to `True`.

---

## Geometry Support

GhDucky can write Rhino geometry to DuckDB's `GEOMETRY` type via WKB encoding and read it back. The spatial extension is loaded using the **Ducky Enable Spatial** component.

### Supported Geometry (Write → WKB)

| Rhino Type | WKB Type |
|---|---|
| `Point3d` / `Point` | POINT Z |
| `LineCurve` / `PolylineCurve` / `Polyline` | LINESTRING Z |
| Other `Curve` | LINESTRING Z (tessellated) |
| `Mesh` | MULTIPOLYGON Z (one polygon per face) |
| `Brep` | MULTIPOLYGON Z (auto-meshed) |

### Supported Geometry (WKB → Read)

| WKB Type | Rhino Type |
|---|---|
| POINT | `Point` |
| LINESTRING | `PolylineCurve` |
| POLYGON | `Mesh` (outer ring; holes ignored with warning) |
| MULTIPOLYGON | `Mesh` (combined) |
| MULTIPOINT | `PointCloud` (all points preserved) |
| MULTILINESTRING | `PolylineCurve` (segments combined with warning) |

---

## Project Structure

```
GhDucky/
├── Components/
│   ├── DuckyComponentBase.cs      # Shared base class
│   ├── Connect/                   # Connect, Disconnect, Spatial, Excel
│   ├── Import/                    # CSV, JSON, Parquet, Excel, DataTree, Geometry
│   ├── Query/                     # Inspect, Query, QueryGeometry, QueryBuilder, Filter, Join
│   ├── Export/                    # Export, ExportExcel
│   └── Excel/                     # ExcelRead, ExcelWrite, ExcelClear, ExcelTemplate, ExcelInfo
├── Services/
│   ├── DuckDBConnectionManager.cs # Process-wide session registry
│   ├── DuckDBSession.cs           # Thread-safe connection wrapper
│   ├── NativeLibraryResolver.cs   # Resolves native DuckDB library for Grasshopper
│   ├── SpatialExtension.cs        # Spatial extension facade
│   ├── ExcelExtension.cs          # Excel extension facade (depends on spatial)
│   ├── DuckDbExtensionTracker.cs  # Generic extension install/load/clear tracker
│   ├── ExcelReader.cs             # Range reads via MiniExcel + contiguous-region expansion
│   ├── ExcelWriter.cs             # In-place cell writes and clears via the OpenXML SDK
│   ├── ExcelTemplate.cs           # Placeholder templating via MiniExcel
│   ├── ExcelReadCache.cs          # Solve-to-solve parsed-range cache
│   ├── ExcelWorkbookMetadata.cs   # Defined names + Excel Tables read from the xlsx package
│   └── ExcelFileGuard.cs          # Lock / extension detection for clear error messages
├── Utils/
│   ├── DuckyColumnType.cs         # SQL column type enum
│   ├── DuckDbAppenderFactory.cs   # Typed per-column appender delegates
│   ├── ExcelAddress.cs            # A1 notation parsing / formatting
│   ├── ExceptionFormatter.cs      # User-friendly exception messages
│   ├── GH_DuckDBConnection.cs     # Grasshopper Goo wrapper for sessions
│   ├── IconFactory.cs             # Emoji-based component icon generator
│   ├── SqlIdentifier.cs           # SQL identifier quoting / injection prevention
│   ├── TypeMapping.cs             # CLR ↔ DuckDB ↔ Grasshopper type conversion
│   └── WkbCodec.cs                # ISO WKB encoder/decoder for Rhino geometry
├── Parameters/
│   └── ParamDuckyDBConnection.cs  # Custom Grasshopper parameter type
├── GhDuckyPriority.cs             # Assembly priority / cleanup on unload
└── GhDuckyInfo.cs                 # Assembly metadata
```

---

## Two ways to use Excel

Ducky deliberately ships two separate Excel paths, because they solve different problems:

| | **Import/Export Excel** (`2 · Import`, `4 · Export`) | **Excel Read/Write** (`5 · Excel`) |
|---|---|---|
| Needs a database connection | Yes | No |
| Granularity | Whole sheet / whole table | Any cell or range |
| Writes | Regenerate the file | In-place, non-destructive |
| Best for | Getting bulk data in and out of SQL | Driving and reading spreadsheets cell by cell |

**Ducky Import Excel Range** bridges the two: it uses the `5 · Excel` range reader to load any range, defined name or Excel Table straight into a DuckDB table, so you can `SELECT` over a sub-range of a spreadsheet.

---

## Dependencies

| Package | Version        | Licence | Purpose |
|---|----------------|---|---|
| [Grasshopper](https://www.nuget.org/packages/Grasshopper/) | 8.0.23304.9001 | — | Rhino/Grasshopper SDK |
| [DuckDB.NET.Data.Full](https://www.nuget.org/packages/DuckDB.NET.Data.Full/) | 1.5.3          | MIT | DuckDB ADO.NET provider + native binaries |
| [System.Drawing.Common](https://www.nuget.org/packages/System.Drawing.Common/) | 8.0.27         | MIT | Icon rendering (GDI+) |
| [MiniExcel.OpenXml](https://www.nuget.org/packages/MiniExcel.OpenXml/) | 2.0.0-preview.4 | Apache-2.0 | Streaming xlsx **reads** — A1 ranges, open-ended ranges, sheet dimensions, native cell typing |
| [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/) | 3.1.0 | MIT | xlsx **writes** — anchored, style-preserving in-place cell edits |

All runtime dependencies are MIT or Apache-2.0, and therefore compatible with this project's MIT licence.

> **Why two Excel libraries?** MiniExcel is excellent at reading but has no cell-level write — its only write verbs replace an entire sheet or file, which destroys formatting, formulas and neighbouring data. The OpenXML SDK provides the in-place write. See [ADR 2](ADR.md).

---

## Tests

Pure-logic unit tests live in `GhDucky.Tests/` (xUnit, net8.0). They cover
`SqlIdentifier`, the non-Grasshopper parts of `TypeMapping`, the WKB encoder in
`WkbCodec`, `ExcelAddress` parsing/formatting, and an end-to-end Excel
round-trip suite that writes real workbooks to a temp folder and verifies that
in-place writes preserve neighbouring cells, formulas and other sheets. Run
from the repo root:

```bash
dotnet test Ducky.slnx
```

> **Note:** `dotnet test` rebuilds `GhDucky` with `EnableDynamicLoading=false`,
> which leaves `GhDucky/bin/` without its dependency DLLs. Run
> `dotnet build GhDucky/GhDucky.csproj` afterwards before copying the folder to
> your Grasshopper libraries directory.

See `GhDucky.Tests/README.md` for what's intentionally out of scope (decode
paths and component-level integration require a running Rhino host).

---

## Contributing

Contributions are welcome! Please open an issue or pull request on [GitHub](https://github.com/mitchell-tesch/Ducky).

---

## License

Copyright © GhDucky Contributors. See [LICENSE](LICENSE) for details.

## Attributions

Ducky Icon: <a href="https://www.flaticon.com/free-icons/rubber-duck" title="rubber duck icons">Rubber duck icons created by Talha Dogar - Flaticon</a>
