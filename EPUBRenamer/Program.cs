// EPUBRenamer (Program.cs)
// A console tool that proposes and applies safe, consistent filenames for .epub files
// based on embedded metadata (Title, Authors). It previews changes by default and can
// copy or move files to a separate output directory when --apply is specified.
//
// Usage:
//   dotnet run -- <inputFolder> [--apply] [--move] [--out <outputFolder>] [--recursive] [--ascii [true|false]]
//
// Options:
//   --apply       Perform the operations (copy/move). Without this flag, runs as a preview.
//   --move        Move files instead of copying (copy is the safer default).
//   --out         Destination folder. Defaults to <inputFolder>\Renamed_<yyyyMMdd_HHmm>.
//   --recursive   Include subdirectories when scanning for .epub files.
//   --ascii       Strip diacritics to improve ASCII-only compatibility (Unicode kept by default).
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
using System.Text;
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

        var proposedMappings = new List<RenameItem>();
        foreach (var epubPath in epubFiles)
        {
            try
            {
                // Read EPUB metadata (title/authors). VersOne.Epub supports EPUB 2/3.
                var book = await EpubReader.ReadBookAsync(epubPath);
                var title = SafeTrim(book?.Title);
                var authors = GetAuthors(book);

                // Fallback to filename when title is missing
                if (string.IsNullOrWhiteSpace(title))
                {
                    title = Path.GetFileNameWithoutExtension(epubPath);
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
        string? outputDirectory = null;
        bool apply = false;
        bool move = false;
        bool recursive = false;
        bool ascii = false; // Unicode by default

        for (int i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg.StartsWith("--"))
            {
                // Handle switches and flags with optional values
                switch (arg)
                {
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
            Ascii = ascii
        };
    }

    /// <summary>
    /// Prints usage information and default behaviors.
    /// </summary>
    private static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  dotnet run -- <inputFolder> [--apply] [--move] [--out <outputFolder>] [--recursive] [--ascii [true|false]]");
        Console.WriteLine();
        Console.WriteLine("Defaults:");
        Console.WriteLine("  - Unicode filenames (no ASCII stripping).");
        Console.WriteLine("  - Top-level search only (non-recursive).");
        Console.WriteLine("  - Copy to <input>\\Renamed_<yyyyMMdd_HHmm> unless --out is provided.");
        Console.WriteLine("  - Author joiner: comma \", \".");
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
        cleaned = cleaned.TrimEnd('.', ' ');

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
