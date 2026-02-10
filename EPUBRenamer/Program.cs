// EPUBRenamer (Program.cs)
// A console tool that proposes and applies safe, consistent filenames for .epub files
// based on embedded metadata (Title, Authors). It previews changes by default and can
// copy or move files to a separate output directory when --apply is specified.
//
// Inspect mode intent:
//   This tool can also run in a report-only "inspection" mode to help you review a large
//   library of EPUBs and identify which files likely warrant a closer look.
//
//   The goal is to flag EPUBs that appear to have been post-processed by common editing or
//   conversion tools (for example, Calibre or Sigil) after the original publication package
//   was created. Such post-processing is often well-intentioned, but it can introduce
//   formatting quirks or non-publisher artifacts, and you may prefer to replace those EPUBs
//   with higher-quality sources.
//
//   The inspection report is heuristic: it searches for known tool fingerprints inside the
//   EPUB container (OPF/XHTML/CSS). A match suggests "this was likely modified," not that the
//   content is necessarily wrong or unsafe.
//
// Usage:
//   dotnet run -- <inputFolder> [--apply] [--move] [--out <outputFolder>] [--recursive] [--ascii [true|false]] [--titlecase [true|false]]
//   dotnet run -- --inspect <folder> [--recursive]
//
// Options:
//   --inspect     Inspect EPUBs and report tool fingerprints (report-only; incompatible with rename/apply options).
//   --apply       Perform the operations (copy/move). Without this flag, runs as a preview.
//   --move        Move files instead of copying (copy is the safer default).
//   --out         Destination folder. Defaults to <inputFolder>\Renamed_<yyyyMMdd_HHmm>.
//   --recursive   Include subdirectories when scanning for .epub files.
//   --ascii       Strip diacritics to improve ASCII-only compatibility (Unicode kept by default).
//   --titlecase   Apply smart Title Case to title and author parts (on by default).
//   --authorformat Set author name formatting: as-is | firstlast | lastfirst (default: firstlast).
//
// Defaults:
//   - Unicode preserved in filenames; punctuation normalized to safe characters.
//   - Author joiner: ", " (comma + space).
//   - Non-recursive search.
//   - Copy to a timestamped output folder (no overwrites; collisions get suffixes).
//   - CSV log (rename-log.csv) is written to the output folder when applying.
//
// Safety:
//   - No in-place overwrites; output is isolated in a new/existing directory.
//   - Collisions are resolved by appending " (1)", " (2)", ... to filenames.
//   - Unreadable/corrupt EPUBs are skipped with warnings.
//
// Dependency:
//   - VersOne.Epub: reads EPUB 2/3 metadata (Title, Authors). This tool does not alter EPUB contents.

using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using VersOne.Epub;

/// <summary>
/// Entry point for the console application.
/// </summary>
/// <param name="args">Command line arguments.</param>
/// <returns>Process exit code.</returns>
return await EpubRenamerApp.RunAsync(args);

/// <summary>
/// Main application containing CLI parsing, EPUB metadata reading,
/// filename generation, preview, and apply logic.
/// </summary>
internal static class EpubRenamerApp
{
    /// <summary>
    /// Cached set of invalid filename characters for the current platform.
    /// Used when sanitizing generated filenames.
    /// </summary>
    static readonly char[] InvalidFileNameChars = Path.GetInvalidFileNameChars();

    /// <summary>
    /// Parses options, enumerates EPUB files, reads metadata, builds proposed
    /// filenames, previews the plan, and (optionally) applies copy/move.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <returns>Exit code (0 on success).</returns>
    public static async Task<int> RunAsync(string[] args)
    {
        // Parse CLI arguments and validate basic invariants
        var options = ParseArguments(args);
        if (options == null)
        {
            PrintUsage();
            return 1;
        }

        // Validate input directory exists to avoid confusing errors later
        if (!Directory.Exists(options.InputDirectory))
        {
            Console.Error.WriteLine($"Input directory not found: {options.InputDirectory}");
            return 2;
        }

        // Gather .epub files, either top-level or recursively
        var searchOption = options.Recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        var epubFiles = Directory.EnumerateFiles(options.InputDirectory, "*.epub", searchOption).ToList();
        if (epubFiles.Count == 0)
        {
            Console.WriteLine("No .epub files found.");
            return 0;
        }

        if (options.Inspect)
        {
            RunInspectionReport(epubFiles);
            return 0;
        }

        var proposedMappings = new List<RenameItem>();
        var missingAuthorFiles = new List<string>();
        var missingTitleFiles = new List<string>();
        var inferredAuthorSummaries = new List<string>();
        var authorFormatChangeSummaries = new List<string>();
        var titleCaseChangeSummaries = new List<string>();
        foreach (var epubPath in epubFiles)
        {
            try
            {
                // Read EPUB metadata (title/authors). VersOne.Epub supports EPUB 2/3.
                var book = await EpubReader.ReadBookAsync(epubPath);
                var rawTitle = SafeTrim(book?.Title);
                var title = rawTitle;
                var authors = GetAuthors(book);

                // Track missing metadata from EPUB (before fallbacks)
                if (string.IsNullOrWhiteSpace(rawTitle))
                {
                    missingTitleFiles.Add(Path.GetFileName(epubPath));
                }
                bool missingAuthorsInMetadata = authors.Count == 0;

                // If author metadata is missing, try to infer from filename (conservative)
                if (missingAuthorsInMetadata)
                {
                    var baseName = Path.GetFileNameWithoutExtension(epubPath);
                    if (TryInferAuthorFromFileName(baseName, out var inferredAuthor))
                    {
                        authors.Add(inferredAuthor);
                        inferredAuthorSummaries.Add($"File: {Path.GetFileName(epubPath)} | Inferred author from filename: \"{Abbrev(inferredAuthor)}\"");
                    }
                    else
                    {
                        missingAuthorFiles.Add(Path.GetFileName(epubPath));
                    }
                }

                // Normalize author formatting (if requested) before titlecasing/joining
                if (authors.Count > 0 && options.AuthorFormat != AuthorFormat.AsIs)
                {
                    var authorsBeforeFormat = authors.ToList();
                    for (int ai = 0; ai < authors.Count; ai++)
                    {
                        authors[ai] = NormalizeAuthor(authors[ai], options.AuthorFormat);
                    }

                    var beforeJoined = string.Join(", ", authorsBeforeFormat);
                    var afterJoined = string.Join(", ", authors);
                    if (!string.Equals(beforeJoined, afterJoined, StringComparison.Ordinal))
                    {
                        var fileOnly = Path.GetFileName(epubPath);
                        authorFormatChangeSummaries.Add($"File: {fileOnly} | AuthorFormat Authors: \"{Abbrev(beforeJoined)}\" -> \"{Abbrev(afterJoined)}\"");
                    }
                }

                // Fallback to filename when title is missing
                if (string.IsNullOrWhiteSpace(title))
                {
                    title = Path.GetFileNameWithoutExtension(epubPath);
                }

                // Optionally Title Case the title and authors before building filename
                if (options.TitleCase)
                {
                    var preTitle = title!;
                    var authorsBeforeTc = authors.ToList();
                    var tcTitle = TitleCaseSmart(preTitle, TitleCaseKind.Title);
                    bool changed = !string.Equals(preTitle, tcTitle, StringComparison.Ordinal);

                    if (authors.Count > 0)
                    {
                        bool anyAuthorChanged = false;
                        for (int ai = 0; ai < authors.Count; ai++)
                        {
                            var before = authors[ai];
                            var after = TitleCaseSmart(before, TitleCaseKind.Author);
                            if (!string.Equals(before, after, StringComparison.Ordinal))
                            {
                                anyAuthorChanged = true;
                            }
                            authors[ai] = after;
                        }
                        changed = changed || anyAuthorChanged;
                    }

                    if (changed)
                    {
                        string? titleBefore = !string.Equals(preTitle, tcTitle, StringComparison.Ordinal) ? preTitle : null;
                        string? titleAfter = titleBefore != null ? tcTitle : null;

                        string? authorsBefore = null;
                        string? authorsAfter = null;
                        if (authorsBeforeTc.Count > 0)
                        {
                            var beforeJoined = string.Join(", ", authorsBeforeTc);
                            var afterJoined = string.Join(", ", authors);
                            if (!string.Equals(beforeJoined, afterJoined, StringComparison.Ordinal))
                            {
                                authorsBefore = beforeJoined;
                                authorsAfter = afterJoined;
                            }
                        }

                        var fileOnly = Path.GetFileName(epubPath);
                        var pieces = new List<string>();
                        if (titleBefore != null)
                        {
                            pieces.Add($"TitleCase Title: \"{Abbrev(titleBefore)}\" -> \"{Abbrev(titleAfter!)}\"");
                        }
                        if (authorsBefore != null)
                        {
                            pieces.Add($"TitleCase Authors: \"{Abbrev(authorsBefore)}\" -> \"{Abbrev(authorsAfter!)}\"");
                        }
                        if (pieces.Count > 0)
                        {
                            titleCaseChangeSummaries.Add($"File: {fileOnly} | {string.Join("; ", pieces)}");
                        }
                    }

                    title = tcTitle;
                }

                // Join multiple authors using a comma+space for Send-to-Kindle friendly naming
                var authorsJoined = authors.Count > 0 ? string.Join(", ", authors) : "Unknown Author";

                // Build sanitized filename stem and append extension
                var proposedFileNameWithoutExt = BuildProposedFileStem(title!, authorsJoined, options.Ascii);
                var finalFileName = EnsureValidFileName($"{proposedFileNameWithoutExt}.epub", options.Ascii);

                proposedMappings.Add(new RenameItem(epubPath, finalFileName));
            }
            catch (Exception ex)
            {
                // Be resilient: skip unreadable/corrupt files and continue
                Console.Error.WriteLine($"[WARN] Skipping unreadable EPUB: {epubPath} ({ex.GetType().Name}: {ex.Message})");
            }
        }

        if (proposedMappings.Count == 0)
        {
            Console.WriteLine("No readable EPUBs found.");
            return 0;
        }

        // Determine output directory (default to Renamed_<timestamp> under input)
        var outputDirectory = options.OutputDirectory;
        if (string.IsNullOrWhiteSpace(outputDirectory))
        {
            outputDirectory = Path.Combine(options.InputDirectory, $"Renamed_{DateTime.Now:yyyyMMdd_HHmm}");
        }

        // Resolve collisions in target directory scope to avoid overwrites.
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in proposedMappings)
        {
            item.TargetFileName = ResolveCollision(item.TargetFileName, usedNames, outputDirectory);
            usedNames.Add(item.TargetFileName);
        }

        // Preview the plan: show Original => Proposed without making changes
        Console.WriteLine();
        Console.WriteLine("Planned renames (preview):");
        Console.WriteLine(new string('-', 80));
        foreach (var item in proposedMappings)
        {
            var originalName = Path.GetFileName(item.SourcePath);
            Console.WriteLine($"{originalName}  =>  {item.TargetFileName}");
        }
        Console.WriteLine(new string('-', 80));
        Console.WriteLine($"Total: {proposedMappings.Count} file(s)");
        Console.WriteLine();

        // If user did not request --apply, we exit after the preview
        if (!options.Apply)
        {
            // Summaries for dry-run to highlight metadata gaps and titlecase impacts
            if (missingAuthorFiles.Count > 0
                || missingTitleFiles.Count > 0
                || inferredAuthorSummaries.Count > 0
                || (options.AuthorFormat != AuthorFormat.AsIs && authorFormatChangeSummaries.Count > 0)
                || (options.TitleCase && titleCaseChangeSummaries.Count > 0))
            {
                Console.WriteLine("Summary:");
                Console.WriteLine(new string('-', 80));
                bool wroteAnySection = false;
                if (missingAuthorFiles.Count > 0)
                {
                    if (wroteAnySection) Console.WriteLine();
                    Console.WriteLine($"Missing author ({missingAuthorFiles.Count}):");
                    foreach (var name in missingAuthorFiles)
                    {
                        Console.WriteLine($"  - {name}");
                    }
                    wroteAnySection = true;
                }
                if (missingTitleFiles.Count > 0)
                {
                    if (wroteAnySection) Console.WriteLine();
                    Console.WriteLine($"Missing title ({missingTitleFiles.Count}):");
                    foreach (var name in missingTitleFiles)
                    {
                        Console.WriteLine($"  - {name}");
                    }
                    wroteAnySection = true;
                }
                if (inferredAuthorSummaries.Count > 0)
                {
                    if (wroteAnySection) Console.WriteLine();
                    Console.WriteLine($"Author inferred from filename ({inferredAuthorSummaries.Count}):");
                    foreach (var line in inferredAuthorSummaries)
                    {
                        Console.WriteLine($"  - {line}");
                    }
                    wroteAnySection = true;
                }
                if (options.AuthorFormat != AuthorFormat.AsIs && authorFormatChangeSummaries.Count > 0)
                {
                    if (wroteAnySection) Console.WriteLine();
                    Console.WriteLine($"Authorformat-adjusted ({authorFormatChangeSummaries.Count}):");
                    foreach (var line in authorFormatChangeSummaries)
                    {
                        Console.WriteLine($"  - {line}");
                    }
                    wroteAnySection = true;
                }
                if (options.TitleCase && titleCaseChangeSummaries.Count > 0)
                {
                    if (wroteAnySection) Console.WriteLine();
                    Console.WriteLine($"Titlecase-adjusted ({titleCaseChangeSummaries.Count}):");
                    foreach (var line in titleCaseChangeSummaries)
                    {
                        Console.WriteLine($"  - {line}");
                    }
                    wroteAnySection = true;
                }
                Console.WriteLine(new string('-', 80));
                Console.WriteLine();
            }
            Console.WriteLine("Dry run only. Re-run with --apply to perform copy/move.");
            return 0;
        }

        // Ensure output directory exists before writing any files
        Directory.CreateDirectory(outputDirectory);

        // CSV log helps with traceability and recovery
        var logLines = new List<string> { "Original,New" };
        int copied = 0, moved = 0, skipped = 0;
        foreach (var item in proposedMappings)
        {
            try
            {
                var destinationPath = Path.Combine(outputDirectory, item.TargetFileName);
                if (options.Move)
                {
                    // Move to safe destination; no overwrite (collisions already handled)
                    File.Move(item.SourcePath, destinationPath);
                    moved++;
                }
                else
                {
                    // Copy by default for safety (preserves originals)
                    File.Copy(item.SourcePath, destinationPath, overwrite: false);
                    copied++;
                }
                logLines.Add($"{CsvEscape(item.SourcePath)},{CsvEscape(destinationPath)}");
            }
            catch (IOException ioEx)
            {
                // Log I/O-specific issues (e.g., permission denied, file exists, path too long)
                Console.Error.WriteLine($"[WARN] Skipped (I/O): {item.SourcePath} -> {item.TargetFileName} ({ioEx.Message})");
                skipped++;
            }
            catch (Exception ex)
            {
                // Log unexpected issues and continue with others
                Console.Error.WriteLine($"[WARN] Skipped: {item.SourcePath} -> {item.TargetFileName} ({ex.GetType().Name}: {ex.Message})");
                skipped++;
            }
        }

        // Persist the mapping for audit and potential rollback/verification
        var logPath = Path.Combine(outputDirectory, "rename-log.csv");
        File.WriteAllLines(logPath, logLines, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

        Console.WriteLine();
        Console.WriteLine($"Output directory: {outputDirectory}");
        Console.WriteLine($"Copied: {copied}, Moved: {moved}, Skipped: {skipped}");
        Console.WriteLine($"Log written: {logPath}");

        return 0;
    }

    /// <summary>
    /// Parses command-line arguments into an <see cref="Options"/> object.
    /// </summary>
    /// <param name="args">Raw command line arguments.</param>
    /// <returns>Options or null if parsing failed.</returns>
    private static Options? ParseArguments(string[] args)
    {
        if (args.Length == 0)
        {
            return null;
        }

        string? inputDirectory = null;
        string? inspectDirectory = null;
        string? outputDirectory = null;
        bool apply = false;
        bool move = false;
        bool recursive = false;
        bool ascii = false; // Unicode by default
        bool titleCase = true; // On by default
        AuthorFormat authorFormat = AuthorFormat.FirstLast; // Default to First Last

        for (int i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg.StartsWith("--"))
            {
                // Handle switches and flags with optional values
                switch (arg)
                {
                    case "--inspect":
                        if (i + 1 >= args.Length)
                        {
                            Console.Error.WriteLine("Missing value for --inspect");
                            return null;
                        }
                        inspectDirectory = args[++i];
                        break;
                    case "--apply":
                        apply = true;
                        break;
                    case "--move":
                        move = true;
                        break;
                    case "--recursive":
                        recursive = true;
                        break;
                    case "--ascii":
                        ascii = true;
                        // Optional explicit bool: --ascii true/false (default true if value omitted)
                        if (i + 1 < args.Length && !args[i + 1].StartsWith("-"))
                        {
                            if (TryParseBool(args[i + 1], out var parsed))
                            {
                                ascii = parsed;
                                i++;
                            }
                        }
                        break;
                    case "--titlecase":
                        titleCase = true;
                        // Optional explicit bool: --titlecase true/false (default true if value omitted)
                        if (i + 1 < args.Length && !args[i + 1].StartsWith("-"))
                        {
                            if (TryParseBool(args[i + 1], out var parsedTc))
                            {
                                titleCase = parsedTc;
                                i++;
                            }
                        }
                        break;
                    case "--authorformat":
                        if (i + 1 >= args.Length)
                        {
                            Console.Error.WriteLine("Missing value for --authorformat (expected: as-is | firstlast | lastfirst)");
                            return null;
                        }
                        var fmt = args[++i];
                        if (!TryParseAuthorFormat(fmt, out authorFormat))
                        {
                            Console.Error.WriteLine($"Invalid value for --authorformat: {fmt} (expected: as-is | firstlast | lastfirst)");
                            return null;
                        }
                        break;
                    case "--out":
                        if (i + 1 >= args.Length)
                        {
                            Console.Error.WriteLine("Missing value for --out");
                            return null;
                        }
                        outputDirectory = args[++i];
                        break;
                    default:
                        Console.Error.WriteLine($"Unknown option: {arg}");
                        return null;
                }
            }
            else
            {
                // First non-option is the input directory
                if (inputDirectory == null)
                {
                    // Capture positional input directory
                    inputDirectory = arg;
                }
                else
                {
                    Console.Error.WriteLine($"Unexpected argument: {arg}");
                    return null;
                }
            }
        }

        if (!string.IsNullOrWhiteSpace(inspectDirectory))
        {
            if (!string.IsNullOrWhiteSpace(inputDirectory))
            {
                Console.Error.WriteLine("Do not provide an input folder when using --inspect. Use: --inspect <folder>");
                return null;
            }
            inputDirectory = inspectDirectory;

            // Reject incompatible flags for inspect mode
            if (apply || move || !string.IsNullOrWhiteSpace(outputDirectory) || ascii || !titleCase || authorFormat != AuthorFormat.FirstLast)
            {
                Console.Error.WriteLine("--inspect is report-only and cannot be combined with rename/apply options.");
                return null;
            }
        }

        if (string.IsNullOrWhiteSpace(inputDirectory))
        {
            return null;
        }

        return new Options
        {
            InputDirectory = inputDirectory,
            OutputDirectory = outputDirectory,
            Apply = apply,
            Move = move,
            Recursive = recursive,
            Ascii = ascii,
            TitleCase = titleCase,
            AuthorFormat = authorFormat,
            Inspect = !string.IsNullOrWhiteSpace(inspectDirectory)
        };
    }

    /// <summary>
    /// Prints usage information and default behaviors.
    /// </summary>
    private static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  dotnet run -- <inputFolder> [--apply] [--move] [--out <outputFolder>] [--recursive] [--ascii [true|false]] [--titlecase [true|false]] [--authorformat <as-is|firstlast|lastfirst>]");
        Console.WriteLine("  dotnet run -- --inspect <folder> [--recursive]");
        Console.WriteLine();
        Console.WriteLine("Defaults:");
        Console.WriteLine("  - Unicode filenames (no ASCII stripping).");
        Console.WriteLine("  - Top-level search only (non-recursive).");
        Console.WriteLine("  - Copy to <input>\\Renamed_<yyyyMMdd_HHmm> unless --out is provided.");
        Console.WriteLine("  - Author joiner: comma \", \".");
        Console.WriteLine("  - Title Case enabled (use --titlecase false to disable).");
        Console.WriteLine("  - Author format: firstlast (use --authorformat as-is to disable).");
    }

    /// <summary>
    /// A single fingerprint match found while inspecting an EPUB.
    /// </summary>
    private sealed class InspectionFinding
    {
        /// <summary>
        /// The tool or family of tools suggested by the fingerprint (e.g., "Calibre", "Sigil").
        /// </summary>
        public string Tool { get; init; } = string.Empty;

        /// <summary>
        /// A relative weight used for scoring. Higher weights indicate stronger evidence.
        /// </summary>
        public int Weight { get; init; }

        /// <summary>
        /// Where the fingerprint was found (path inside the EPUB archive).
        /// </summary>
        public string Location { get; init; } = string.Empty;

        /// <summary>
        /// A short human-readable description of the evidence (token/attribute detected).
        /// </summary>
        public string Evidence { get; init; } = string.Empty;
    }

    /// <summary>
    /// Summary of inspection results for a single EPUB file.
    /// </summary>
    private sealed class InspectionResult
    {
        /// <summary>
        /// Full filesystem path to the EPUB.
        /// </summary>
        public string EpubPath { get; init; } = string.Empty;

        /// <summary>
        /// Sum of all finding weights. This is a heuristic score, not a probability.
        /// </summary>
        public int Score { get; init; }

        /// <summary>
        /// Individual evidence items discovered during inspection.
        /// </summary>
        public List<InspectionFinding> Findings { get; init; } = new();

        /// <summary>
        /// If non-null, inspection failed for this EPUB (corrupt ZIP, missing OPF, etc.).
        /// </summary>
        public string? Error { get; init; }
    }

    /// <summary>
    /// Runs an inspection pass and prints a human-readable report.
    ///
    /// Output is intentionally optimized for quick scanning:
    /// - A compact "All files" list (alphabetical) showing OK/SUSPECT/ERROR.
    /// - A second "Suspicious files" list (alphabetical) that includes the top reasons.
    /// </summary>
    private static void RunInspectionReport(List<string> epubFiles)
    {
        // Chosen to be conservative: require at least one high-confidence fingerprint.
        const int suspiciousThreshold = 5;

        var results = new List<InspectionResult>(capacity: epubFiles.Count);
        foreach (var epubPath in epubFiles)
        {
            results.Add(InspectEpub(epubPath));
        }

        results.Sort((a, b) => StringComparer.OrdinalIgnoreCase.Compare(Path.GetFileName(a.EpubPath), Path.GetFileName(b.EpubPath)));

        Console.WriteLine();
        Console.WriteLine("Inspection report (heuristic):");
        Console.WriteLine(new string('-', 80));
        Console.WriteLine($"Scanned: {results.Count} file(s)");
        Console.WriteLine($"Flagged (suspicious): {results.Count(r => r.Error != null || r.Score >= suspiciousThreshold)}");
        Console.WriteLine(new string('-', 80));
        Console.WriteLine();

        Console.WriteLine("All files (alphabetical):");
        foreach (var r in results)
        {
            var name = Path.GetFileName(r.EpubPath);
            if (r.Error != null)
            {
                Console.WriteLine($"{name}  [ERROR] {r.Error}");
            }
            else if (r.Score >= suspiciousThreshold)
            {
                Console.WriteLine($"{name}  [SUSPECT] (score={r.Score})");
            }
            else
            {
                Console.WriteLine($"{name}  [OK]");
            }
        }

        var suspicious = results
            .Where(r => r.Error != null || r.Score >= suspiciousThreshold)
            .ToList();

        Console.WriteLine();
        Console.WriteLine(new string('-', 80));
        Console.WriteLine($"Suspicious files (alphabetical): {suspicious.Count}");
        Console.WriteLine(new string('-', 80));

        foreach (var r in suspicious)
        {
            var name = Path.GetFileName(r.EpubPath);
            if (r.Error != null)
            {
                Console.WriteLine();
                Console.WriteLine($"{name}  [ERROR]");
                Console.WriteLine($"  - {r.Error}");
                continue;
            }

            Console.WriteLine();
            Console.WriteLine($"{name}  (score={r.Score})");
            foreach (var f in r.Findings
                .OrderByDescending(f => f.Weight)
                .ThenBy(f => f.Tool, StringComparer.OrdinalIgnoreCase)
                .Take(6))
            {
                Console.WriteLine($"  - {f.Tool}: {f.Evidence} ({f.Location})");
            }
        }

        var toolBreakdown = suspicious
            .SelectMany(r => r.Findings.Select(f => f.Tool).Distinct(StringComparer.OrdinalIgnoreCase))
            .GroupBy(t => t, StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(g => g.Count())
            .ThenBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (toolBreakdown.Count > 0)
        {
            Console.WriteLine();
            Console.WriteLine(new string('-', 80));
            Console.WriteLine("Breakdown (flagged files mentioning tool):");
            foreach (var g in toolBreakdown)
            {
                Console.WriteLine($"  {g.Key}: {g.Count()}");
            }
        }

        Console.WriteLine();
        Console.WriteLine("Note: This is heuristic. A fingerprint suggests post-processing, not necessarily “bad” edits.");
    }

    /// <summary>
    /// Inspects a single EPUB (ZIP container) and returns a set of fingerprint findings.
    ///
    /// This routine is read-only: it does not extract or modify the EPUB. It reads only a
    /// small subset of internal files needed for fingerprints to keep bulk scans fast.
    /// </summary>
    private static InspectionResult InspectEpub(string epubPath)
    {
        var findings = new List<InspectionFinding>();
        try
        {
            using var archive = ZipFile.OpenRead(epubPath);

            // container.xml points to the OPF package document (EPUB 2/3).
            string? containerPath = FindEntryPath(archive, "META-INF/container.xml");
            if (containerPath == null)
            {
                return new InspectionResult { EpubPath = epubPath, Error = "Missing META-INF/container.xml" };
            }

            var containerText = ReadEntryTextLimited(archive, containerPath, maxBytes: 256 * 1024);
            var opfPath = TryGetOpfPathFromContainer(containerText);
            if (string.IsNullOrWhiteSpace(opfPath))
            {
                return new InspectionResult { EpubPath = epubPath, Error = "Could not locate OPF path from container.xml" };
            }

            string? opfEntryPath = FindEntryPath(archive, opfPath);
            if (opfEntryPath == null)
            {
                return new InspectionResult { EpubPath = epubPath, Error = $"OPF not found in archive: {opfPath}" };
            }

            var opfText = ReadEntryTextLimited(archive, opfEntryPath, maxBytes: 1024 * 1024);
            AnalyzeOpf(opfText, opfEntryPath, findings);

            // Scan XHTML/CSS (capped) with priority for titlepage-like files.
            //
            // Why cap?
            //   Some EPUBs contain hundreds of resources. We want inspection to remain fast and
            //   avoid reading megabytes of content per book. Many fingerprints occur in a small
            //   set of common files (OPF, calibre.css, titlepage.xhtml), so we bias our scan
            //   toward those first.
            const int maxEntriesToScan = 30;
            const int maxBytesPerEntry = 256 * 1024;
            var candidateEntries = archive.Entries
                .Where(e => !string.IsNullOrWhiteSpace(e.FullName))
                .Where(e =>
                {
                    var ext = Path.GetExtension(e.FullName).ToLowerInvariant();
                    return ext is ".xhtml" or ".html" or ".htm" or ".css";
                })
                .Where(e =>
                    !e.FullName.Equals(containerPath, StringComparison.OrdinalIgnoreCase) &&
                    !e.FullName.Equals(opfEntryPath, StringComparison.OrdinalIgnoreCase))
                .ToList();

            var prioritized = candidateEntries
                .Select(e => new { Entry = e, Priority = GetInspectPriority(e.FullName) })
                .OrderBy(x => x.Priority)
                .ThenBy(x => x.Entry.FullName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var toScan = new List<ZipArchiveEntry>(capacity: Math.Min(maxEntriesToScan, prioritized.Count));

            // Always include "titlepage" priority first; then fill remaining slots.
            foreach (var x in prioritized)
            {
                if (x.Priority == 0)
                {
                    toScan.Add(x.Entry);
                }
            }
            foreach (var x in prioritized)
            {
                if (toScan.Count >= maxEntriesToScan)
                {
                    break;
                }
                if (!toScan.Any(e => e.FullName.Equals(x.Entry.FullName, StringComparison.OrdinalIgnoreCase)))
                {
                    toScan.Add(x.Entry);
                }
            }

            foreach (var entry in toScan)
            {
                var text = ReadEntryTextLimited(archive, entry.FullName, maxBytesPerEntry);
                AnalyzeContentText(text, entry.FullName, findings);
            }

            // Current scoring model: simple weighted sum of evidence items.
            // This is intentionally easy to reason about and tweak as new fingerprints are added.
            int score = findings.Sum(f => f.Weight);
            return new InspectionResult { EpubPath = epubPath, Score = score, Findings = findings };
        }
        catch (InvalidDataException ex)
        {
            return new InspectionResult { EpubPath = epubPath, Error = $"Invalid EPUB/ZIP: {ex.Message}" };
        }
        catch (Exception ex)
        {
            return new InspectionResult { EpubPath = epubPath, Error = $"{ex.GetType().Name}: {ex.Message}" };
        }
    }

    private static int GetInspectPriority(string entryName)
    {
        // Lower numeric value = scanned earlier.
        // We always prioritize "titlepage" variants because Calibre/Sigil signatures are often present there.
        var n = entryName.ToLowerInvariant();
        if (n.Contains("titlepage") || n.Contains("title-page") || n.Contains("title_page"))
        {
            return 0;
        }
        if (n.Contains("calibre") || n.Contains("cover") || n.Contains("nav") || n.Contains("toc"))
        {
            return 1;
        }
        return 2;
    }

    private static string? FindEntryPath(ZipArchive archive, string desiredPath)
    {
        // Most EPUBs are case-sensitive inside the ZIP, but Windows is not. Be tolerant here.
        var direct = archive.GetEntry(desiredPath);
        if (direct != null)
        {
            return direct.FullName;
        }
        // Case-insensitive fallback
        foreach (var e in archive.Entries)
        {
            if (e.FullName.Equals(desiredPath, StringComparison.OrdinalIgnoreCase))
            {
                return e.FullName;
            }
        }
        return null;
    }

    private static string ReadEntryTextLimited(ZipArchive archive, string entryPath, int maxBytes)
    {
        // Reads only the first maxBytes bytes to keep scans fast; fingerprints are typically near the top.
        var entry = archive.GetEntry(entryPath) ?? archive.Entries.First(e => e.FullName.Equals(entryPath, StringComparison.OrdinalIgnoreCase));
        using var stream = entry.Open();
        using var ms = new MemoryStream();
        var buffer = new byte[16 * 1024];
        int remaining = maxBytes;
        while (remaining > 0)
        {
            int read = stream.Read(buffer, 0, Math.Min(buffer.Length, remaining));
            if (read <= 0)
            {
                break;
            }
            ms.Write(buffer, 0, read);
            remaining -= read;
        }
        // Decode as UTF-8 with fallback; we only need rough text scanning.
        return Encoding.UTF8.GetString(ms.ToArray());
    }

    private static string? TryGetOpfPathFromContainer(string containerXml)
    {
        try
        {
            // container.xml is small, so a simple XML parse is fine.
            var doc = XDocument.Parse(containerXml);
            var rootfile = doc
                .Descendants()
                .FirstOrDefault(e => e.Name.LocalName.Equals("rootfile", StringComparison.OrdinalIgnoreCase));
            return rootfile?.Attribute("full-path")?.Value;
        }
        catch
        {
            // Fallback: regex extract full-path="..."
            var m = Regex.Match(containerXml, "full-path\\s*=\\s*\"(?<p>[^\"]+)\"", RegexOptions.IgnoreCase);
            return m.Success ? m.Groups["p"].Value : null;
        }
    }

    private static void AnalyzeOpf(string opfText, string opfLocation, List<InspectionFinding> findings)
    {
        var lower = opfText.ToLowerInvariant();

        // Calibre fingerprints.
        // Calibre commonly writes "calibre:"-prefixed metadata and may include its namespace URL.
        if (lower.Contains("https://calibre-ebook.com") || lower.Contains("calibre-ebook.com"))
        {
            findings.Add(new InspectionFinding { Tool = "Calibre", Weight = 5, Location = opfLocation, Evidence = "OPF contains calibre-ebook.com prefix/URL" });
        }

        string[] calibreTokens =
        {
            "calibre:series", "calibre:series_index", "calibre:title_sort", "calibre:timestamp", "calibre:user_metadata", "calibre:author_link_map"
        };
        foreach (var tok in calibreTokens)
        {
            if (lower.Contains(tok))
            {
                findings.Add(new InspectionFinding { Tool = "Calibre", Weight = 6, Location = opfLocation, Evidence = $"OPF contains {tok}" });
                break;
            }
        }

        // Generator fingerprints (common place for tool identification).
        foreach (Match m in Regex.Matches(opfText, "name\\s*=\\s*\"generator\"[^>]*content\\s*=\\s*\"(?<g>[^\"]+)\"", RegexOptions.IgnoreCase))
        {
            var g = m.Groups["g"].Value.Trim();
            if (g.Length == 0) continue;
            var gl = g.ToLowerInvariant();
            if (gl.Contains("sigil"))
            {
                findings.Add(new InspectionFinding { Tool = "Sigil", Weight = 6, Location = opfLocation, Evidence = $"generator=\"{g}\"" });
            }
            else if (gl.Contains("pandoc"))
            {
                findings.Add(new InspectionFinding { Tool = "Pandoc", Weight = 6, Location = opfLocation, Evidence = $"generator=\"{g}\"" });
            }
            else if (gl.Contains("kindlegen"))
            {
                findings.Add(new InspectionFinding { Tool = "KindleGen", Weight = 6, Location = opfLocation, Evidence = $"generator=\"{g}\"" });
            }
            else if (gl.Contains("kindle"))
            {
                findings.Add(new InspectionFinding { Tool = "Kindle", Weight = 4, Location = opfLocation, Evidence = $"generator=\"{g}\"" });
            }
            else if (gl.Contains("adobe") || gl.Contains("indesign"))
            {
                findings.Add(new InspectionFinding { Tool = "InDesign/Adobe", Weight = 3, Location = opfLocation, Evidence = $"generator=\"{g}\"" });
            }
        }
    }

    private static void AnalyzeContentText(string text, string location, List<InspectionFinding> findings)
    {
        var lower = text.ToLowerInvariant();

        // Calibre XHTML/CSS fingerprints.
        // Calibre's conversion/templates frequently include a "calibre" class/id/selector.
        if (lower.Contains("class=\"calibre\"") || lower.Contains("class='calibre'") ||
            lower.Contains("id=\"calibre") || lower.Contains("id='calibre") ||
            lower.Contains(".calibre") || lower.Contains("#calibre"))
        {
            findings.Add(new InspectionFinding { Tool = "Calibre", Weight = 5, Location = location, Evidence = "XHTML/CSS contains calibre class/id/selector" });
        }

        // Calibre can also stamp XHTML with calibre-specific <meta name="calibre:..."> markers,
        // commonly used to identify cover/title pages (e.g., calibre:cover).
        if (lower.Contains("name=\"calibre:") || lower.Contains("name='calibre:"))
        {
            // Prefer a more specific hint when we can see a well-known token.
            if (lower.Contains("name=\"calibre:cover\"") || lower.Contains("name='calibre:cover'"))
            {
                findings.Add(new InspectionFinding { Tool = "Calibre", Weight = 5, Location = location, Evidence = "XHTML contains <meta name=\"calibre:cover\" ...>" });
            }
            else
            {
                findings.Add(new InspectionFinding { Tool = "Calibre", Weight = 4, Location = location, Evidence = "XHTML contains <meta name=\"calibre:*\" ...>" });
            }
        }
    }

    /// <summary>
    /// Extracts author names from the <see cref="EpubBook"/> in a version-tolerant way.
    /// Supports multiple possible property names across VersOne.Epub versions.
    /// </summary>
    /// <param name="book">Loaded EPUB book or null.</param>
    /// <returns>List of author names; may be empty.</returns>
    private static List<string> GetAuthors(EpubBook? book)
    {
        var authors = new List<string>();
        try
        {
            if (book == null)
            {
                return authors;
            }

            // Prefer list when available (older versions exposed AuthorList)
            var listFromProperty = (book.GetType().GetProperty("AuthorList")?.GetValue(book) as IEnumerable<string>)?.ToList();
            if (listFromProperty != null && listFromProperty.Count > 0)
            {
                return listFromProperty.Select(a => SafeTrim(a)).Where(a => !string.IsNullOrWhiteSpace(a)).ToList()!;
            }

            // Try Authors (newer API exposes Authors as IEnumerable<string>)
            var authorsProp = book.GetType().GetProperty("Authors")?.GetValue(book) as IEnumerable<string>;
            if (authorsProp != null)
            {
                authors = authorsProp.Select(a => SafeTrim(a)).Where(a => !string.IsNullOrWhiteSpace(a)).ToList()!;
                if (authors.Count > 0)
                {
                    return authors;
                }
            }

            // Fallback to single Author string (some versions provide a single Author)
            var author = SafeTrim(book.GetType().GetProperty("Author")?.GetValue(book) as string);
            if (!string.IsNullOrWhiteSpace(author))
            {
                authors.Add(author!);
            }
        }
        catch
        {
            // Be permissive; return whatever we could get without failing the whole run
        }
        return authors;
    }

    /// <summary>
    /// Generates the filename stem using the requested pattern and sanitization rules.
    /// </summary>
    /// <param name="title">Book title (already trimmed or fallback).</param>
    /// <param name="authorsJoined">Authors joined with ", ".</param>
    /// <param name="ascii">True to strip diacritics and normalize to ASCII-friendly text.</param>
    /// <returns>Sanitized filename stem (without extension).</returns>
    private static string BuildProposedFileStem(string title, string authorsJoined, bool ascii)
    {
        var titlePart = EnsureValidFileName(title, ascii);
        var authorsPart = EnsureValidFileName(authorsJoined, ascii);

        var stem = $"{titlePart} - {authorsPart}";
        stem = CollapseWhitespace(stem);
        stem = stem.Trim();

        // Limit to keep room for extension
        const int maxStemLength = 120;
        if (stem.Length > maxStemLength)
        {
            stem = stem.Substring(0, maxStemLength).TrimEnd('.', ' ');
        }
        return stem;
    }

    /// <summary>
    /// Normalizes and sanitizes a candidate filename or segment by replacing
    /// invalid characters, collapsing whitespace, and avoiding reserved names.
    /// Optionally strips diacritics for ASCII-only compatibility.
    /// </summary>
    /// <param name="input">Raw input text.</param>
    /// <param name="ascii">Whether to strip diacritics (ASCII mode).</param>
    /// <returns>Safe filename string.</returns>
    private static string EnsureValidFileName(string input, bool ascii)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return "Untitled";
        }

        string normalized = NormalizePunctuation(input);
        if (ascii)
        {
            normalized = RemoveDiacritics(normalized);
        }

        // Replace invalid filename chars
        var builder = new StringBuilder(normalized.Length);
        foreach (var ch in normalized)
        {
            if (InvalidFileNameChars.Contains(ch))
            {
                builder.Append(' ');
            }
            else
            {
                builder.Append(ch);
            }
        }

        var cleaned = CollapseWhitespace(builder.ToString()).Trim();
        // Avoid trailing punctuation that often appears in scraped metadata (e.g., "Tolhurst;").
        cleaned = cleaned.TrimEnd('.', ' ', ';', ':', ',');

        if (string.IsNullOrWhiteSpace(cleaned))
        {
            return "Untitled";
        }

        // Avoid reserved device names (Windows)
        var reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "CON","PRN","AUX","NUL",
            "COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9",
            "LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9"
        };
        if (reserved.Contains(cleaned))
        {
            cleaned = $"_{cleaned}";
        }

        return cleaned;
    }

    /// <summary>
    /// Ensures a unique target filename by appending numeric suffixes when
    /// another file with the same name is already planned or exists on disk.
    /// </summary>
    /// <param name="fileName">Proposed filename (with extension).</param>
    /// <param name="usedNames">Set of already-reserved names in this run.</param>
    /// <param name="outputDirectory">Directory to check for existing files.</param>
    /// <returns>Unique filename.</returns>
    private static string ResolveCollision(string fileName, HashSet<string> usedNames, string outputDirectory)
    {
        string baseName = Path.GetFileNameWithoutExtension(fileName);
        string ext = Path.GetExtension(fileName);
        string candidate = fileName;
        int index = 1;
        while (usedNames.Contains(candidate) || File.Exists(Path.Combine(outputDirectory, candidate)))
        {
            candidate = $"{baseName} ({index}){ext}";
            index++;
        }
        return candidate;
    }

    /// <summary>
    /// Replaces common typographic punctuation with ASCII-friendly equivalents.
    /// This improves compatibility while preserving readability.
    /// </summary>
    /// <param name="input">Input text.</param>
    /// <returns>Text with punctuation normalized.</returns>
    private static string NormalizePunctuation(string input)
    {
        // Replace common typographic punctuation with ASCII equivalents
        return input
            .Replace('—', '-')  // em dash
            .Replace('–', '-')  // en dash
            .Replace('―', '-')  // horizontal bar
            .Replace('…', '.')  // will collapse to "..." through invalid-char/space handling + collapsing
            .Replace('“', '"')
            .Replace('”', '"')
            .Replace('‘', '\'')
            .Replace('’', '\'');
    }

    /// <summary>
    /// Removes diacritics (accents) from text by decomposing the string
    /// and discarding combining marks, then re-composing to Form C.
    /// </summary>
    /// <param name="text">Input text.</param>
    /// <returns>ASCII-friendly version of the text.</returns>
    private static string RemoveDiacritics(string text)
    {
        var normalized = text.Normalize(NormalizationForm.FormD);
        var builder = new StringBuilder(capacity: normalized.Length);
        foreach (var ch in normalized)
        {
            var uc = CharUnicodeInfo.GetUnicodeCategory(ch);
            if (uc != UnicodeCategory.NonSpacingMark && uc != UnicodeCategory.SpacingCombiningMark && uc != UnicodeCategory.EnclosingMark)
            {
                builder.Append(ch);
            }
        }
        return builder.ToString().Normalize(NormalizationForm.FormC);
    }

    /// <summary>
    /// Collapses any run of whitespace characters into a single space and trims
    /// sequences to a cleaner, filesystem-friendly form.
    /// </summary>
    /// <param name="input">Text potentially containing extra whitespace.</param>
    /// <returns>Text with normalized whitespace.</returns>
    private static string CollapseWhitespace(string input)
    {
        var builder = new StringBuilder(input.Length);
        bool previousWasSpace = false;
        foreach (var ch in input)
        {
            if (char.IsWhiteSpace(ch))
            {
                if (!previousWasSpace)
                {
                    builder.Append(' ');
                    previousWasSpace = true;
                }
            }
            else
            {
                builder.Append(ch);
                previousWasSpace = false;
            }
        }
        return builder.ToString();
    }

    /// <summary>
    /// Trims a string safely, handling null inputs by returning null.
    /// </summary>
    /// <param name="s">Input string or null.</param>
    /// <returns>Trimmed string or null.</returns>
    private static string? SafeTrim(string? s) => s?.Trim();

    /// <summary>
    /// Parses a lenient boolean from common tokens like 1/0, true/false, yes/no, y/n.
    /// </summary>
    /// <param name="value">Raw token.</param>
    /// <param name="result">Parsed result when successful.</param>
    /// <returns>True if parsing succeeded.</returns>
    private static bool TryParseBool(string value, out bool result)
    {
        switch (value.Trim().ToLowerInvariant())
        {
            case "1":
            case "true":
            case "yes":
            case "y":
                result = true;
                return true;
            case "0":
            case "false":
            case "no":
            case "n":
                result = false;
                return true;
            default:
                result = false;
                return false;
        }
    }

    private enum AuthorFormat
    {
        AsIs,
        FirstLast,
        LastFirst
    }

    private static bool TryParseAuthorFormat(string value, out AuthorFormat format)
    {
        switch (value.Trim().ToLowerInvariant())
        {
            case "as-is":
            case "asis":
            case "as_is":
                format = AuthorFormat.AsIs;
                return true;
            case "firstlast":
            case "first-last":
            case "first_last":
                format = AuthorFormat.FirstLast;
                return true;
            case "lastfirst":
            case "last-first":
            case "last_first":
                format = AuthorFormat.LastFirst;
                return true;
            default:
                format = AuthorFormat.AsIs;
                return false;
        }
    }

    /// <summary>
    /// Normalizes an author string with conservative rules:
    /// - Cleans spacing around commas and collapses whitespace.
    /// - If a comma is present, treats "Last, First [Middle]" as reorderable.
    /// - If no comma is present, leaves order as-is (avoids guessing).
    /// </summary>
    private static string NormalizeAuthor(string author, AuthorFormat format)
    {
        if (string.IsNullOrWhiteSpace(author))
        {
            return author;
        }

        var s = CollapseWhitespace(author.Trim());

        // Normalize comma spacing: "Doe ,  Jane" => "Doe, Jane"
        s = s.Replace(" ,", ",").Replace(", ", ", ").Replace(",,", ",");
        while (s.Contains("  "))
        {
            s = s.Replace("  ", " ");
        }

        // Fix missing spaces after initials in author fields even when not titlecasing
        // e.g. "J.Kent Layton" -> "J. Kent Layton"
        s = InsertSpaceAfterInitialPeriod(s);

        if (format == AuthorFormat.AsIs)
        {
            return s;
        }

        // Conservative reorder: only if comma exists.
        int commaIdx = s.IndexOf(',');
        if (commaIdx < 0)
        {
            return s;
        }

        var parts = s.Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(p => CollapseWhitespace(p.Trim()))
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .ToList();

        if (parts.Count < 2)
        {
            return s;
        }

        // If there are too many comma-separated segments, it's likely multiple authors in a single string.
        // Example: "Foo Bar, Zoo Goo, Baz Qux" — do not attempt to reorder.
        if (parts.Count > 3)
        {
            return s;
        }

        // Handle common suffixes if present as third component: "Doe, Jane, Jr."
        string last = parts[0];
        string firstMiddle = parts[1];
        string? suffix = null;
        if (parts.Count >= 3 && IsLikelySuffix(parts[2]))
        {
            suffix = parts[2];
        }
        else if (parts.Count == 3)
        {
            // 3 segments but third isn't a suffix; likely multiple authors or extra qualifiers.
            return s;
        }

        // Guardrail: avoid reordering when both sides look like full names (likely multiple authors in one string).
        // Example: "Foo Bar, Zoo Goo" should be treated as "Foo Bar" + "Zoo Goo" (two authors), not "Last, First".
        // We don't split authors here; we just avoid making it worse.
        int lastWords = CollapseWhitespace(last).Split(' ', StringSplitOptions.RemoveEmptyEntries).Length;
        int firstWords = CollapseWhitespace(firstMiddle).Split(' ', StringSplitOptions.RemoveEmptyEntries).Length;
        if (lastWords >= 2 && firstWords >= 2)
        {
            return s;
        }

        if (format == AuthorFormat.FirstLast)
        {
            var result = suffix == null
                ? $"{firstMiddle} {last}"
                : $"{firstMiddle} {last} {suffix}";
            return CollapseWhitespace(result).Trim();
        }
        else // LastFirst
        {
            var result = suffix == null
                ? $"{last}, {firstMiddle}"
                : $"{last}, {firstMiddle}, {suffix}";
            return CollapseWhitespace(result).Trim();
        }
    }

    private static bool IsLikelySuffix(string value)
    {
        var v = value.Trim().Trim('.').ToLowerInvariant();
        return v is "jr" or "sr" or "ii" or "iii" or "iv" or "v";
    }

    /// <summary>
    /// Escapes a CSV field by quoting and doubling any embedded quotes when necessary.
    /// </summary>
    /// <param name="value">Raw value.</param>
    /// <returns>Escaped CSV field.</returns>
    private static string CsvEscape(string value)
    {
        if (value.Contains('"') || value.Contains(',') || value.Contains('\n') || value.Contains('\r'))
        {
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
        return value;
    }

    /// <summary>
    /// Abbreviates a string to a maximum length with an ellipsis when necessary.
    /// </summary>
    /// <param name="value">Input string.</param>
    /// <param name="maxLength">Maximum allowed length.</param>
    /// <returns>Possibly abbreviated string.</returns>
    private static string Abbrev(string value, int maxLength = 80)
    {
        if (string.IsNullOrEmpty(value) || value.Length <= maxLength)
        {
            return value;
        }
        if (maxLength <= 3)
        {
            return value.Substring(0, Math.Max(0, maxLength));
        }
        return value.Substring(0, maxLength - 3) + "...";
    }

    /// <summary>
    /// Applies smart Title Case to a string:
    /// - Lowercases minor words (and, or, the, a, an, in, on, of, to, at, by, for, from, nor, but, as, per, vs, via)
    ///   unless they are the first or last word or follow a colon.
    /// - Capitalizes the first letter of other words (and letters after apostrophes).
    /// - Capitalizes each segment of hyphenated words.
    /// - Preserves all-caps acronyms (2+ letters).
    /// </summary>
    /// <param name="input">Raw input.</param>
    /// <returns>Title-cased string.</returns>
    private enum TitleCaseKind
    {
        Title,
        Author
    }

    private static string TitleCaseSmart(string input, TitleCaseKind kind)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return input;
        }

        // Minor words to keep lowercase in the middle (style-guide-ish, not perfect)
        var minorWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "and","or","the","a","an","in","on","of","to","at","by","for","from",
            "nor","but","as","per","vs","via","with","into","onto","off","up","down"
        };

        // Author particles like "del" should usually stay lowercase in the middle of names.
        if (kind == TitleCaseKind.Author)
        {
            minorWords.Add("del");
        }

        // Split on whitespace but keep original spacing
        var parts = input.Split(' ', StringSplitOptions.None);
        bool prevEndedWithBoundary = false;
        for (int i = 0; i < parts.Length; i++)
        {
            var token = parts[i];
            if (token.Length == 0)
            {
                continue;
            }

            bool isFirst = i == 0;
            bool isLast = i == parts.Length - 1;

            // Preserve leading/trailing punctuation while transforming core word
            int start = 0;
            int end = token.Length - 1;
            while (start <= end && !char.IsLetterOrDigit(token[start]))
            {
                start++;
            }
            while (end >= start && !char.IsLetterOrDigit(token[end]))
            {
                end--;
            }

            if (start > end)
            {
                // No alphanumerics, leave as-is
                prevEndedWithBoundary = token is ":" or "·" || token.EndsWith(":") || token.EndsWith("·");
                continue;
            }

            string leading = token.Substring(0, start);
            string core = token.Substring(start, end - start + 1);
            string trailing = token.Substring(end + 1);

            // Process hyphenated segments individually
            var hyphenSegments = core.Split('-');
            for (int h = 0; h < hyphenSegments.Length; h++)
            {
                string seg = hyphenSegments[h];
                if (seg.Length == 0)
                {
                    continue;
                }

                // If this token is a single-letter initial like "A.", never treat it as the article "a".
                bool isInitial = seg.Length == 1 && trailing.StartsWith(".", StringComparison.Ordinal);

                bool treatAsMinor = !isInitial && !isFirst && !isLast && !prevEndedWithBoundary && minorWords.Contains(seg);

                if (IsAllCapsAcronym(seg))
                {
                    // Keep acronyms (e.g., NASA, AI) uppercase
                    hyphenSegments[h] = seg.ToUpperInvariant();
                }
                else if (isInitial)
                {
                    hyphenSegments[h] = seg.ToUpperInvariant();
                }
                else if (treatAsMinor)
                {
                    hyphenSegments[h] = seg.ToLowerInvariant();
                }
                else
                {
                    hyphenSegments[h] = CapitalizeWordPreservingApostrophes(seg);
                }
            }

            var transformedCore = string.Join("-", hyphenSegments);
            parts[i] = leading + transformedCore + trailing;

            prevEndedWithBoundary =
                trailing.Contains(":") || trailing.Contains("·")
                // Closing quotes often indicate a new segment starts next (e.g., ..."Candidate" The CIA...)
                || trailing.Contains("\"") || trailing.Contains("”") || trailing.Contains("»")
                || transformedCore.EndsWith(":") || transformedCore.EndsWith("·")
                || token.EndsWith(":") || token.EndsWith("·")
                // If this token ends with a comma and the next token starts with an uppercase letter,
                // treat it as a new title/segment (e.g., "... Narcissus, The Secret Agent").
                || (trailing.Contains(",") && i + 1 < parts.Length && TokenStartsWithUpper(parts[i + 1]));
        }

        return string.Join(" ", parts);
    }

    private static bool TokenStartsWithUpper(string token)
    {
        if (string.IsNullOrWhiteSpace(token))
        {
            return false;
        }

        for (int i = 0; i < token.Length; i++)
        {
            char ch = token[i];
            if (char.IsLetter(ch))
            {
                return char.IsUpper(ch);
            }
        }
        return false;
    }

    private static bool TryInferAuthorFromFileName(string baseName, out string author)
    {
        author = string.Empty;
        if (string.IsNullOrWhiteSpace(baseName))
        {
            return false;
        }

        // Conservative: only infer when the filename clearly ends with " -- Author"
        // Example: "A New Day Yesterday ... -- Mike Barnes"
        const string sep = " -- ";
        int idx = baseName.LastIndexOf(sep, StringComparison.Ordinal);
        if (idx < 0)
        {
            return false;
        }

        var candidate = baseName.Substring(idx + sep.Length).Trim();
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return false;
        }

        // Reject obvious non-author tails (mostly numeric / ids)
        int letters = 0, digits = 0;
        foreach (var ch in candidate)
        {
            if (char.IsLetter(ch)) letters++;
            else if (char.IsDigit(ch)) digits++;
        }
        if (letters < 2 || digits > letters)
        {
            return false;
        }

        // Avoid very long "author" tails (likely metadata dump)
        if (candidate.Length > 60)
        {
            return false;
        }

        author = CollapseWhitespace(candidate);
        return !string.IsNullOrWhiteSpace(author);
    }

    /// <summary>
    /// Fixes patterns like "X.Foo" to "X. Foo" (common in initials) while avoiding
    /// lower-case abbreviations like "e.g." by requiring the next letter be uppercase.
    /// </summary>
    private static string InsertSpaceAfterInitialPeriod(string input)
    {
        var sb = new StringBuilder(input.Length + 4);
        for (int i = 0; i < input.Length; i++)
        {
            char ch = input[i];
            sb.Append(ch);

            if (ch == '.' && i - 1 >= 0 && i + 1 < input.Length)
            {
                char prev = input[i - 1];
                char next = input[i + 1];
                if (char.IsLetter(prev) && char.IsUpper(next))
                {
                    // Don't add if already spaced
                    if (next != ' ')
                    {
                        sb.Append(' ');
                    }
                }
            }
        }
        return sb.ToString();
    }

    private static bool IsAllCapsAcronym(string word)
    {
        int letters = 0;
        foreach (var ch in word)
        {
            if (char.IsLetter(ch))
            {
                letters++;
                if (!char.IsUpper(ch))
                {
                    return false;
                }
            }
        }
        return letters >= 2;
    }

    private static bool LooksIntentionallyCased(string word)
    {
        // Preserve casing for things like "McCarthy", "iPhone", "eBay" where internal capitals are meaningful.
        int firstLetterIndex = -1;
        for (int i = 0; i < word.Length; i++)
        {
            if (char.IsLetter(word[i]))
            {
                firstLetterIndex = i;
                break;
            }
        }
        if (firstLetterIndex < 0)
        {
            return false;
        }

        // If there's an uppercase letter after the first letter, treat as intentional casing.
        for (int i = firstLetterIndex + 1; i < word.Length; i++)
        {
            if (char.IsUpper(word[i]))
            {
                return true;
            }
        }
        return false;
    }

    private static string CapitalizeWordPreservingApostrophes(string word)
    {
        if (LooksIntentionallyCased(word))
        {
            return word;
        }

        // Lowercase everything then uppercase first letter and any letter after an apostrophe
        var lower = word.ToLowerInvariant();
        var builder = new StringBuilder(lower.Length);
        bool capitalizeNext = true;
        int lettersSinceStart = 0;
        char? firstLetterUpper = null;
        for (int i = 0; i < lower.Length; i++)
        {
            char ch = lower[i];
            if (capitalizeNext && char.IsLetter(ch))
            {
                builder.Append(char.ToUpperInvariant(ch));
                capitalizeNext = false;
            }
            else
            {
                builder.Append(ch);
            }

            if (char.IsLetter(ch))
            {
                lettersSinceStart++;
                firstLetterUpper ??= char.ToUpperInvariant(ch);
            }

            // Capitalize after apostrophes ONLY for common single-letter prefixes (O'Brien, D'Artagnan, L'Étranger),
            // not for possessives/contractions (Hitchhiker's, don't, it's).
            if ((ch == '\'' || ch == '’') && i + 1 < lower.Length && char.IsLetter(lower[i + 1]))
            {
                if (lettersSinceStart == 1 && firstLetterUpper is 'O' or 'D' or 'L')
                {
                    capitalizeNext = true;
                }
            }
        }

        var result = builder.ToString();

        // Handle common "Mc" prefix: "mccarthy" -> "McCarthy"
        if (result.Length >= 3 && result.StartsWith("Mc", StringComparison.OrdinalIgnoreCase) && char.IsLetter(result[2]))
        {
            result = result.Substring(0, 2) + char.ToUpperInvariant(result[2]) + result.Substring(3);
        }

        return result;
    }

    /// <summary>
    /// Parsed options model for a single app invocation.
    /// </summary>
    private sealed class Options
    {
        /// <summary>
        /// Folder to scan for .epub files.
        /// </summary>
        public string InputDirectory { get; set; } = string.Empty;

        /// <summary>
        /// Destination folder for renamed files (created if missing).
        /// Defaults to &lt;input&gt;\Renamed_&lt;timestamp&gt; when not specified.
        /// </summary>
        public string? OutputDirectory { get; set; }

        /// <summary>
        /// When true, performs copy/move; otherwise, preview only.
        /// </summary>
        public bool Apply { get; set; }

        /// <summary>
        /// When true, moves files instead of copying (riskier).
        /// </summary>
        public bool Move { get; set; }

        /// <summary>
        /// When true, scans subdirectories recursively.
        /// </summary>
        public bool Recursive { get; set; }

        /// <summary>
        /// When true, strips diacritics to increase ASCII compatibility.
        /// </summary>
        public bool Ascii { get; set; }

        /// <summary>
        /// When true, applies smart Title Case to title and author fields.
        /// </summary>
        public bool TitleCase { get; set; }

        /// <summary>
        /// Author name formatting behavior.
        /// </summary>
        public AuthorFormat AuthorFormat { get; set; } = AuthorFormat.AsIs;

        /// <summary>
        /// When true, runs in inspection/report-only mode (no rename/copy/move).
        /// </summary>
        public bool Inspect { get; set; }
    }

    /// <summary>
    /// Represents a proposed rename operation from a source path to a sanitized target filename.
    /// </summary>
    private sealed class RenameItem
    {
        /// <summary>
        /// Creates a new rename item with the original path and proposed target filename.
        /// </summary>
        /// <param name="sourcePath">Full path to the original file.</param>
        /// <param name="targetFileName">Filename (no directory) for the destination.</param>
        public RenameItem(string sourcePath, string targetFileName)
        {
            SourcePath = sourcePath;
            TargetFileName = targetFileName;
        }

        /// <summary>
        /// Full path to the original file.
        /// </summary>
        public string SourcePath { get; }

        /// <summary>
        /// Proposed filename for the destination file (unique, sanitized).
        /// </summary>
        public string TargetFileName { get; set; }
    }
}
