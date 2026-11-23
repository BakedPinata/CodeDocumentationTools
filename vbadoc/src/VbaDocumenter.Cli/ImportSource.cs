using CodeDocumenter.Shared;
using CodeDocumenter.Workbooks.Importing;
using Serilog;
using VbaDocumenter.Core.Analysis;
using VbaDocumenter.Core.Extraction;
using VbaDocumenter.Core.Importing;
using VbaDocumenter.Core.Parsing;

namespace VbaDocumenter.Cli;

/// <summary>What a workbook or folder has that <c>vbadoc import</c> could take, for the desktop app to offer as a list.</summary>
public static class ImportSource
{
  /// <summary>Every module in the workbook, with why it can't be taken if it can't, by the rules <c>import</c> uses.</summary>
  /// <exception cref="ExtractionException">The workbook can't be read, or has no VBA project.</exception>
  public static IReadOnlyList<ImportableItem> Items(string workbook) => Items(workbook, recursive: false);

  /// <summary>
  /// Every module in a workbook, or in a folder of module files (its subfolders too with <paramref name="recursive"/>),
  /// with why it can't be taken if it can't, by the rules <c>import</c> uses.
  /// </summary>
  /// <exception cref="ExtractionException">The source can't be read, or has no modules.</exception>
  public static IReadOnlyList<ImportableItem> Items(string source, bool recursive)
  {
    ArgumentException.ThrowIfNullOrWhiteSpace(source);
    var project = Read(Path.GetFullPath(source), recursive, null);
    return ModuleImport.For(project).Items();
  }

  /// <summary>
  /// The project <c>import</c> takes from: a workbook's, or a folder's as <see cref="ModuleFolderReader.ForImport"/> gives
  /// it, with what reading the folder found to warn about logged.
  /// </summary>
  internal static VbaProject Read(string source, bool recursive, ILogger? logger)
  {
    if (!Directory.Exists(source)) return WorkbookReader.Read(source);
    var folder = ModuleFolderReader.Read(source, recursive, folderGroups: false);
    foreach (var warning in folder.Warnings) logger?.Warning("{Warning:l}", warning);
    return ModuleFolderReader.ForImport(folder.Project);
  }
}
