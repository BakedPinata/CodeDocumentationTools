namespace CodeDocumenter.Shared.Files;

/// <summary>A source file found in a folder.</summary>
/// <param name="FullPath">Its full path.</param>
/// <param name="RelativePath">
/// Its path below the folder read, with <c>/</c> between the parts, such as <c>Reports/Sales.m</c>: what the page keeps,
/// so Load folder finds it again below the folder picked. For a file at the top level it's the file name.
/// </param>
public sealed record SourceFolderFile(string FullPath, string RelativePath)
{
  /// <summary>The subfolders it's in, from the top: empty for a file at the top level.</summary>
  public IReadOnlyList<string> Folders => RelativePath.Split('/')[..^1];
}

/// <summary>Finds the source files in a folder a tool reads, at its top level or, when asked, in its subfolders too.</summary>
public static class SourceFolder
{
  /// <summary>
  /// Folders never read for source, even with <c>recursive</c>: version control's own and build outputs. Hidden
  /// folders, and those starting with a dot, are skipped as well.
  /// </summary>
  public static IReadOnlySet<string> SkippedFolders { get; } =
    new HashSet<string>(["bin", "obj", "node_modules"], StringComparer.OrdinalIgnoreCase);

  /// <summary>
  /// The files in <paramref name="folder"/> with one of <paramref name="extensions"/> (with their dots, any case),
  /// in order of their relative paths, those at the top level first.
  /// </summary>
  /// <exception cref="DirectoryNotFoundException">The folder does not exist.</exception>
  /// <exception cref="UnauthorizedAccessException">The folder cannot be listed.</exception>
  /// <exception cref="IOException">The folder cannot be listed.</exception>
  public static IReadOnlyList<SourceFolderFile> Find(string folder, IReadOnlyCollection<string> extensions, bool recursive)
  {
    ArgumentException.ThrowIfNullOrWhiteSpace(folder);
    ArgumentNullException.ThrowIfNull(extensions);
    var root = new DirectoryInfo(folder);
    if (!root.Exists)
    {
      throw new DirectoryNotFoundException($"Could not find the folder '{folder}'.");
    }

    var found = new List<SourceFolderFile>();
    Collect(root, string.Empty, extensions, recursive, found);
    return found
      .OrderBy(f => f.RelativePath.Contains('/', StringComparison.Ordinal))
      .ThenBy(f => f.RelativePath, StringComparer.OrdinalIgnoreCase)
      .ToList();
  }

  /// <summary>Whether a subfolder is left out of a recursive read.</summary>
  public static bool IsSkipped(DirectoryInfo directory)
  {
    ArgumentNullException.ThrowIfNull(directory);
    return directory.Name.StartsWith('.')
      || SkippedFolders.Contains(directory.Name)
      || (directory.Attributes & (FileAttributes.Hidden | FileAttributes.System)) != 0;
  }

  private static void Collect(DirectoryInfo directory, string prefix, IReadOnlyCollection<string> extensions, bool recursive, List<SourceFolderFile> found)
  {
    // Hidden and system files and folders are skipped, as they always have been at the top level.
    var options = new EnumerationOptions { RecurseSubdirectories = false, IgnoreInaccessible = false };
    foreach (var file in directory.EnumerateFiles("*", options))
    {
      if (extensions.Any(e => string.Equals(file.Extension, e, StringComparison.OrdinalIgnoreCase)))
      {
        found.Add(new SourceFolderFile(file.FullName, prefix + file.Name));
      }
    }

    if (!recursive)
    {
      return;
    }

    foreach (var sub in directory.EnumerateDirectories("*", options))
    {
      if (!IsSkipped(sub))
      {
        Collect(sub, prefix + sub.Name + "/", extensions, recursive, found);
      }
    }
  }
}
