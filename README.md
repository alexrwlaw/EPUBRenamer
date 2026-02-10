# EPUBRenamer

`EPUBRenamer` is a Windows-friendly console utility for bulk-processing `.epub` files:

- **Rename mode**: proposes consistent, safe filenames based on EPUB metadata (Title + Authors) and can optionally copy/move the files to an output folder.
- **Inspect mode**: scans EPUB internals for common post-processing tool fingerprints (e.g., Calibre/Sigil) and produces a report to help you decide which files warrant review or replacement from a better source.

## Features

- **Preview-first workflow**: shows planned renames before making changes.
- **Safe apply**: when applying, writes to an output directory and avoids overwrites (collisions get numeric suffixes).
- **Smart formatting**:
  - Title Case (enabled by default; can be disabled)
  - Author normalization (default: `Last, First` → `First Last`)
  - Conservative author inference from filename when metadata is missing (for certain filename patterns)
- **Inspect/report mode**:
  - Read-only ZIP scanning (no extraction, no modification)
  - Heuristic scoring + clear reasons
  - Output optimized for humans: an “all files” list and a “suspicious files” list

## Requirements

- .NET SDK (the solution currently targets modern .NET; build output will indicate the exact version)

## Build

From the solution folder:

```powershell
dotnet build .\EPUBRenamer.sln
```

## Usage

### Rename mode (default)

Preview planned renames (no changes made):

```powershell
dotnet run -- .\MyEpubFolder
```

Apply (copy by default) to an output directory:

```powershell
dotnet run -- .\MyEpubFolder --apply
```

Move instead of copy:

```powershell
dotnet run -- .\MyEpubFolder --apply --move
```

Recursive scan:

```powershell
dotnet run -- .\MyEpubFolder --recursive
```

Disable Title Case (Title Case is enabled by default):

```powershell
dotnet run -- .\MyEpubFolder --titlecase false
```

Control author formatting:

```powershell
dotnet run -- .\MyEpubFolder --authorformat as-is
dotnet run -- .\MyEpubFolder --authorformat firstlast
dotnet run -- .\MyEpubFolder --authorformat lastfirst
```

### Inspect mode (report-only)

Inspect a folder of EPUBs and print a report (no rename/copy/move performed):

```powershell
dotnet run -- --inspect .\MyEpubFolder
```

Inspect recursively:

```powershell
dotnet run -- --inspect .\MyEpubFolder --recursive
```

Notes:
- `--inspect` is intentionally **incompatible** with apply/rename options. It is a reporting tool to help you identify EPUBs that may have been modified by conversion/editing tools after publication.

