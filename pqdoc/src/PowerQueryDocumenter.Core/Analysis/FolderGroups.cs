using PowerQueryDocumenter.Core.Extraction;

namespace PowerQueryDocumenter.Core.Analysis;

/// <summary>Query groups made from the subfolders a folder's <c>.m</c> files are in (<c>--folder-groups</c>, FR-IN-7).</summary>
public static class FolderGroups
{
  /// <summary>
  /// The group of each file, by its relative path (with <c>/</c> between the parts): one group per subfolder, nested as
  /// the folders are and ordered by name. A file at the top level has none, and isn't in the result.
  /// </summary>
  public static IReadOnlyDictionary<string, QueryGroup> For(IEnumerable<string> relativePaths)
  {
    ArgumentNullException.ThrowIfNull(relativePaths);
    var paths = relativePaths.ToList();
    var folders = paths
      .Select(p => p.Split('/')[..^1])
      .Where(parts => parts.Length > 0)
      .SelectMany(parts => Enumerable.Range(1, parts.Length).Select(n => string.Join('/', parts[..n])))
      .Distinct(StringComparer.OrdinalIgnoreCase)
      .ToList();

    // Each group's order is its place by name among those beside it.
    var groups = new Dictionary<string, QueryGroup>(StringComparer.OrdinalIgnoreCase);
    foreach (var folder in folders.OrderBy(f => f.Count(c => c == '/')).ThenBy(f => f, StringComparer.OrdinalIgnoreCase))
    {
      var cut = folder.LastIndexOf('/');
      var parent = cut < 0 ? null : groups[folder[..cut]];
      var order = groups.Values.Count(g => ReferenceEquals(g.Parent, parent));
      groups[folder] = new QueryGroup(folder[(cut + 1)..], order, null, parent);
    }

    var result = new Dictionary<string, QueryGroup>(StringComparer.Ordinal);
    foreach (var path in paths)
    {
      var cut = path.LastIndexOf('/');
      if (cut > 0)
      {
        result[path] = groups[path[..cut]];
      }
    }

    return result;
  }
}
