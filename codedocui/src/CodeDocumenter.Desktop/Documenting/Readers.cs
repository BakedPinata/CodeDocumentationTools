using CodeDocumenter.Desktop.Running;
using CodeDocumenter.Shared.Files;
using VbaDocumenter.Core.Extraction;

namespace CodeDocumenter.Desktop.Documenting;

/// <summary>
/// Which tool reads what, to offer only the tools that apply. The tools tell a workbook by what's in it, not by its
/// name, so this follows what each tool's help says it reads; a test keeps the two the same.
/// </summary>
internal static class Readers
{
  /// <summary>vbadoc's workbooks, as its <c>&lt;workbook&gt;</c> help lists them.</summary>
  public static IReadOnlyList<string> VbadocExtensions { get; } = [".xlsm", ".xlsb", ".xlam", ".xls", ".xla"];

  /// <summary>pqdoc's workbooks, as its <c>&lt;input&gt;</c> help lists them. It reads a folder of .m files too.</summary>
  public static IReadOnlyList<string> PqdocExtensions { get; } = [".xlsx", ".xlsm", ".xlsb"];

  /// <summary>
  /// Whether a tool reads an input, as typed and tidied. Nothing chosen yet counts as yes, so every tool is offered. For
  /// pqdoc, a name with no extension that isn't a file is taken as a folder.
  /// </summary>
  public static bool Reads(Tool tool, string input)
  {
    ArgumentNullException.ThrowIfNull(input);
    if (input.Length == 0)
    {
      return true;
    }

    // A folder is read by the tool whose files are in it, at any depth, as --recursive would find them.
    if (Directory.Exists(input))
    {
      return HasFiles(input, tool == Tool.Vbadoc ? ModuleFolderReader.Extensions : [".m"]);
    }

    var extension = Path.GetExtension(input);
    return tool switch
    {
      Tool.Vbadoc => VbadocExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase),
      Tool.Pqdoc => PqdocExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase)
        || (extension.Length == 0 && !File.Exists(input)),
      _ => false,
    };
  }

  static bool HasFiles(string folder, IReadOnlyList<string> extensions)
  {
    try
    {
      return SourceFolder.Find(folder, extensions, recursive: true).Count > 0;
    }
    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
    {
      return false;
    }
  }

  /// <summary>Why a tool isn't offered for an input, for the window.</summary>
  public static string WhyNot(Tool tool, string input)
  {
    var extension = Path.GetExtension(input);
    return tool switch
    {
      Tool.Vbadoc when Directory.Exists(input) => "no .bas, .cls or .frm files in this folder or below it",
      Tool.Pqdoc when Directory.Exists(input) => "no .m files in this folder or below it",
      _ when extension.Length > 0 => $"{tool.Name()} doesn't read {extension.ToLowerInvariant()} files",
      _ => $"{tool.Name()} doesn't read this",
    };
  }
}
