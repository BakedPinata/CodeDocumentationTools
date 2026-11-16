using System.Text;
using CodeDocumenter.Shared.Files;
using PowerQueryDocumenter.Core.Model;
using PowerQueryDocumenter.Core.Output;
using Serilog;

namespace PowerQueryDocumenter.Core.Analysis;

/// <summary>The text of one query.</summary>
/// <param name="Name">The query name: for a folder, the .m file's name without its extension and with its escapes reversed.</param>
/// <param name="Content">The M code.</param>
public sealed record QuerySource(string Name, string Content);

/// <summary>The outcome of analysing a folder.</summary>
/// <param name="Graph">The dependency graph.</param>
/// <param name="QueryFileCount">Number of .m files found, including unreadable and duplicate ones.</param>
/// <param name="UnreadableFiles">Files that could not be read. Their queries are in the graph, but without their own references.</param>
/// <param name="DuplicateFiles">Files ignored because another file has the same query name (names are case-insensitive).</param>
/// <param name="Files">The path of each query's file that was read, by query name, for the page's source view.</param>
/// <param name="RelativePaths">
/// Each read file's path below the folder, with <c>/</c> between the parts, by query name: the file name for one at the
/// top level.
/// </param>
/// <param name="ClashingQueries">Queries whose file name another file in another folder shares (FR-IN-6).</param>
public sealed record AnalysisResult(
  DependencyGraph Graph,
  int QueryFileCount,
  IReadOnlyList<string> UnreadableFiles,
  IReadOnlyList<string> DuplicateFiles,
  IReadOnlyDictionary<string, string>? Files = null,
  IReadOnlyDictionary<string, string>? RelativePaths = null,
  IReadOnlySet<string>? ClashingQueries = null);

/// <summary>
/// Reads Power Query .m files and builds the <see cref="DependencyGraph"/> between them.
/// Analysis is separate from exporting: the result is an immutable graph (FR-ANA-7).
/// </summary>
public sealed class DependencyAnalyser
{
  /// <summary>The file extension of exported Power Query queries.</summary>
  public const string QueryFileExtension = ".m";

  private readonly ILogger logger;
  private readonly Func<string, CancellationToken, Task<string>> readFile;

  /// <param name="logger">Destination for progress and diagnostics.</param>
  /// <param name="readFile">Reads a file's text. Defaults to <see cref="ReadQueryFileAsync"/>; replaceable for testing.</param>
  public DependencyAnalyser(ILogger logger, Func<string, CancellationToken, Task<string>>? readFile = null)
  {
    ArgumentNullException.ThrowIfNull(logger);
    this.logger = logger;
    this.readFile = readFile ?? ReadQueryFileAsync;
  }

  /// <summary>
  /// Reads a query file as UTF-8. A byte order mark, if present, takes precedence, so UTF-16 and
  /// UTF-32 files with a BOM are also read correctly (FR-ANA-8).
  /// </summary>
  public static Task<string> ReadQueryFileAsync(string path, CancellationToken cancellationToken) =>
    File.ReadAllTextAsync(path, Encoding.UTF8, cancellationToken);

  /// <summary>
  /// Analyses every .m file directly inside <paramref name="folderPath"/> (subfolders are not scanned).
  /// A file that cannot be read is logged and skipped (FR-ANA-8). If the folder has no .m files the
  /// result has a <see cref="AnalysisResult.QueryFileCount"/> of zero.
  /// </summary>
  /// <exception cref="DirectoryNotFoundException">The folder does not exist.</exception>
  /// <exception cref="UnauthorizedAccessException">The folder cannot be listed.</exception>
  /// <exception cref="IOException">The folder cannot be listed.</exception>
  public Task<AnalysisResult> AnalyseFolderAsync(string folderPath, CancellationToken cancellationToken) =>
    AnalyseFolderAsync(folderPath, recursive: false, cancellationToken);

  /// <summary>
  /// Analyses every .m file in <paramref name="folderPath"/>, and with <paramref name="recursive"/> in its subfolders
  /// too (FR-IN-5), apart from those <see cref="SourceFolder"/> skips. A query is named after its file. When two files
  /// in different folders give the same name, both are kept: one at the top level, or else the first by path, keeps the name, and references
  /// to it go to that one; the other is named by its path, such as <c>Old/Sales</c>, and both are in
  /// <see cref="AnalysisResult.ClashingQueries"/>. Two files in one folder whose names differ only by case can't both
  /// be told apart, so the second is ignored, as before.
  /// </summary>
  public async Task<AnalysisResult> AnalyseFolderAsync(string folderPath, bool recursive, CancellationToken cancellationToken)
  {
    ArgumentException.ThrowIfNullOrWhiteSpace(folderPath);

    var files = SourceFolder.Find(folderPath, [QueryFileExtension], recursive);
    logger.Debug("Found {Count} query file(s) in {Folder}", files.Count, folderPath);

    var pathsByName = new Dictionary<string, string>(DependencyGraph.NameComparer);
    var relative = new Dictionary<string, string>(DependencyGraph.NameComparer);
    var duplicates = new List<string>();
    var clashing = new HashSet<string>(DependencyGraph.NameComparer);
    foreach (var file in files)
    {
      // A file pqdoc wrote from a workbook has escapes in its name (Sales%2FReturns.m): reverse them (FR-QF-2).
      var name = QueryFileNames.QueryName(Path.GetFileNameWithoutExtension(file.FullPath));
      if (pathsByName.TryAdd(name, file.FullPath))
      {
        relative[name] = file.RelativePath;
        continue;
      }

      var existing = relative[name];
      if (string.Equals(FolderOf(existing), FolderOf(file.RelativePath), StringComparison.OrdinalIgnoreCase))
      {
        duplicates.Add(file.FullPath);
        logger.Warning(
          "Ignoring {File}: query {Query} is already defined by {ExistingFile} (query names are not case-sensitive)",
          file.FullPath,
          name,
          pathsByName[name]);
        continue;
      }

      var pathName = string.Join('/', file.Folders.Append(name));
      if (!pathsByName.TryAdd(pathName, file.FullPath))
      {
        duplicates.Add(file.FullPath);
        continue;
      }

      relative[pathName] = file.RelativePath;
      clashing.Add(name);
      clashing.Add(pathName);
      logger.Warning(
        "Query {Query} is in both {ExistingFile} and {File}. Both are shown, the second as {PathName}, and references to {Query} go to {ExistingFile}. A workbook can't hold both, so import won't take either",
        name,
        existing,
        file.RelativePath,
        pathName);
    }

    var sources = new List<QuerySource>();
    var unreadable = new List<string>();
    var read = new Dictionary<string, string>(DependencyGraph.NameComparer);
    foreach (var (name, path) in pathsByName)
    {
      cancellationToken.ThrowIfCancellationRequested();
      try
      {
        sources.Add(new QuerySource(name, await readFile(path, cancellationToken).ConfigureAwait(false)));
        read[name] = path;
      }
      catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or DecoderFallbackException)
      {
        unreadable.Add(path);
        logger.Error(
          "Could not read {File}: {Reason}. Query {Query} is still included, but its own references are missing",
          path,
          ex.Message,
          name);
        logger.Debug(ex, "Read failure detail for {File}", path);
      }
    }

    var graph = Analyse(pathsByName.Keys, sources);
    var readRelative = read.Keys.ToDictionary(n => n, n => relative[n], DependencyGraph.NameComparer);
    return new AnalysisResult(graph, files.Count, unreadable, duplicates, read, readRelative, clashing);
  }

  private static string FolderOf(string relativePath) =>
    relativePath.LastIndexOf('/') is var at and >= 0 ? relativePath[..at] : string.Empty;

  /// <summary>Builds a graph from in-memory query sources.</summary>
  public DependencyGraph Analyse(IEnumerable<QuerySource> sources)
  {
    ArgumentNullException.ThrowIfNull(sources);
    var list = sources.ToList();
    return Analyse(list.Select(s => s.Name), list);
  }

  /// <summary>
  /// Builds a graph. <paramref name="queryNames"/> is the full set of known queries; a query can
  /// be known (and referenced) even when its own source is unavailable.
  /// </summary>
  public DependencyGraph Analyse(IEnumerable<string> queryNames, IEnumerable<QuerySource> sources)
  {
    ArgumentNullException.ThrowIfNull(queryNames);
    ArgumentNullException.ThrowIfNull(sources);

    var sourceList = sources.ToList();
    var knownQueries = new Dictionary<string, string>(DependencyGraph.NameComparer);
    foreach (var name in queryNames.Concat(sourceList.Select(s => s.Name)))
    {
      knownQueries.TryAdd(name, name);
    }

    var edges = new List<DependencyEdge>();
    var lines = new Dictionary<DependencyEdge, IReadOnlyList<int>>();
    var steps = new Dictionary<string, IReadOnlyList<QueryStep>>(DependencyGraph.NameComparer);
    foreach (var source in sourceList)
    {
      var current = knownQueries[source.Name];
      var structure = ReferenceExtractor.Read(MTokeniser.Tokenise(source.Content));
      var references = FindReferences(source.Content, structure, current, knownQueries);
      var queryStepList = StepsOf(source.Content, structure, current, knownQueries);
      if (queryStepList.Count > 0)
      {
        steps[current] = queryStepList;
      }

      if (references.Count > 0)
      {
        logger.Debug("{Query} references {References}", current, references.Keys);
      }

      foreach (var (upstream, at) in references)
      {
        var edge = new DependencyEdge(upstream, current);
        edges.Add(edge);
        lines[edge] = at;
      }
    }

    return DependencyGraph.Create(knownQueries.Values, edges, lines, steps);
  }

  /// <summary>
  /// Returns the known queries that <paramref name="content"/> refers to, each with the lines (from 1, in order) it's
  /// named on.
  /// </summary>
  /// <remarks>
  /// Only identifiers that match a known query name count as references, so built-in functions
  /// and keywords can never produce a false dependency, and no built-in name filter is needed.
  /// That is what lets a query named <c>Source</c>, <c>Table</c> or <c>Value</c> be detected
  /// (FR-ANA-2). Dotted names such as <c>Table.SelectRows</c> are single identifiers, so they do
  /// not match a query named <c>Table</c>.
  /// </remarks>
  private static SortedDictionary<string, IReadOnlyList<int>> FindReferences(
    string content,
    MQueryStructure structure,
    string currentQueryName,
    Dictionary<string, string> knownQueries)
  {
    var referenced = new SortedDictionary<string, SortedSet<int>>(DependencyGraph.NameComparer);
    var lineStarts = LineStarts(content);

    foreach (var (identifier, position) in structure.Uses)
    {
      if (QueryNamed(identifier, currentQueryName, knownQueries) is not { } query)
      {
        continue;
      }

      if (!referenced.TryGetValue(query, out var at))
      {
        referenced[query] = at = [];
      }

      at.Add(LineAt(lineStarts, position));
    }

    var result = new SortedDictionary<string, IReadOnlyList<int>>(DependencyGraph.NameComparer);
    foreach (var (query, at) in referenced)
    {
      result[query] = at.ToArray();
    }

    return result;
  }

  /// <summary>
  /// The known query an identifier refers to, or null: only known query names count, a lower-case M type name
  /// ("as table") only refers to a query spelled exactly the same way, and a query never refers to itself (FR-ANA-4).
  /// </summary>
  private static string? QueryNamed(string identifier, string currentQueryName, Dictionary<string, string> knownQueries)
  {
    if (!knownQueries.TryGetValue(identifier, out var query))
    {
      return null;
    }

    if (MLanguage.IsTypeWord(identifier) && !string.Equals(identifier, query, StringComparison.Ordinal))
    {
      return null;
    }

    return DependencyGraph.NameComparer.Equals(query, currentQueryName) ? null : query;
  }

  /// <summary>
  /// The outermost <c>let</c>'s steps (step E4), each with its line and the queries it uses, in the order first used,
  /// then the <c>in</c> expression, with no name; none for a query without <c>let</c>.
  /// </summary>
  private static List<QueryStep> StepsOf(
    string content,
    MQueryStructure structure,
    string currentQueryName,
    Dictionary<string, string> knownQueries)
  {
    var result = new List<QueryStep>();
    if (structure.Steps.Count == 0)
    {
      return result;
    }

    var lineStarts = LineStarts(content);
    var starts = structure.Steps.Select(s => s.Position).Append(structure.InPosition ?? int.MaxValue).ToArray();
    var names = structure.Steps.Select(s => (string?)s.Name).Append(null).ToArray();
    for (var k = 0; k < starts.Length; k++)
    {
      if (starts[k] == int.MaxValue)
      {
        break; // a let with no in, as when the text stops short
      }

      var end = k + 1 < starts.Length ? starts[k + 1] : int.MaxValue;
      var uses = structure.Uses
        .Where(u => u.Position >= starts[k] && u.Position < end)
        .Select(u => QueryNamed(u.Name, currentQueryName, knownQueries))
        .OfType<string>()
        .Distinct(DependencyGraph.NameComparer)
        .ToArray();
      result.Add(new QueryStep(names[k], LineAt(lineStarts, starts[k]), uses));
    }

    return result;
  }

  /// <summary>Where each line starts. CRLF, LF and CR each end a line, as the page splits them.</summary>
  private static List<int> LineStarts(string text)
  {
    var starts = new List<int> { 0 };
    for (var i = 0; i < text.Length; i++)
    {
      if (text[i] == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
      {
        i++;
      }

      if (text[i] is '\r' or '\n')
      {
        starts.Add(i + 1);
      }
    }

    return starts;
  }

  /// <summary>The line (from 1) of a character position.</summary>
  private static int LineAt(List<int> lineStarts, int position)
  {
    var found = lineStarts.BinarySearch(position);
    return found >= 0 ? found + 1 : ~found;
  }
}
