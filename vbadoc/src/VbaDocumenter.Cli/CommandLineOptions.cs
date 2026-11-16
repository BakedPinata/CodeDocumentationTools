using CodeDocumenter.Shared.Files;
using VbaDocumenter.Cli.Logging;

namespace VbaDocumenter.Cli;

/// <summary>The bound command line: paths are absolute, and the output folder and page name are always set.</summary>
/// <param name="WorkbookPath">The input: a workbook, or a folder of module files.</param>
/// <param name="OutputName">The page's name without ".html": the workbook's or folder's name and "-vba", or <c>--output-name</c>.</param>
/// <param name="OnExists">What to do when the page already exists and would change (<c>--on-exists</c>, <c>--force</c>).</param>
/// <param name="Recursive">With a folder, whether its subfolders are read too (<c>--recursive</c>).</param>
/// <param name="FolderGroups">With <paramref name="Recursive"/>, whether subfolders become explorer folders (<c>--folder-groups</c>).</param>
sealed record CommandLineOptions(
    string WorkbookPath,
    string OutputFolder,
    string OutputName,
    bool EmbedSource,
    OnExistsBehaviour OnExists,
    LogVerbosity Verbosity,
    bool Recursive = false,
    bool FolderGroups = false)
{
  /// <summary>Whether the input is a folder of module files rather than a workbook.</summary>
  public bool InputIsFolder => Directory.Exists(WorkbookPath);
}
