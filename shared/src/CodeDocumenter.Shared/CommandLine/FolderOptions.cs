using System.CommandLine;

namespace CodeDocumenter.Shared.CommandLine;

/// <summary>
/// <c>--recursive</c> and <c>--folder-groups</c>, the same in every tool that reads a folder of source files: whether its
/// subfolders are read too, and whether they become the explorer's groups.
/// </summary>
public sealed class FolderOptions
{
  /// <summary>Creates the options.</summary>
  /// <param name="files">The files read, for the help text, such as ".m files".</param>
  /// <param name="groups">What subfolders become with <c>--folder-groups</c>, such as "query groups".</param>
  public FolderOptions(string files, string groups)
  {
    Recursive = CreateRecursive(files);
    FolderGroups = new Option<bool>("--folder-groups")
    {
      Description = $"With --recursive, show each subfolder as one of the {groups} in the explorer.",
    };
  }

  /// <summary><c>--recursive</c> / <c>-r</c>.</summary>
  public Option<bool> Recursive { get; }

  /// <summary><c>--folder-groups</c>.</summary>
  public Option<bool> FolderGroups { get; }

  /// <summary>A <c>--recursive</c> option on its own, for a command that has no groups, such as <c>import</c>.</summary>
  public static Option<bool> CreateRecursive(string files) => new("--recursive", "-r")
  {
    Description = $"With a folder as input, read the {files} in its subfolders too. Folders starting with a dot, hidden "
      + "folders, and bin, obj and node_modules are skipped.",
  };
}
