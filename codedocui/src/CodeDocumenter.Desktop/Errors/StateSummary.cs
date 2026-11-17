using System.Globalization;
using CodeDocumenter.Desktop.Documenting;
using CodeDocumenter.Desktop.Importing;
using CodeDocumenter.Desktop.Running;

namespace CodeDocumenter.Desktop.Errors;

/// <summary>What the window was showing and set to when an error came.</summary>
/// <param name="Tab">The tab showing.</param>
/// <param name="Running">Whether a run was going.</param>
/// <param name="ScalePercent">The screen's scale, such as 150.</param>
/// <param name="Document">The Document tab.</param>
/// <param name="Import">The Import tab.</param>
/// <param name="Listed">How many rows the Import tab's list has.</param>
/// <param name="Recent">How many recent workbooks and folders there are.</param>
internal sealed record AppState(string Tab, bool Running, int ScalePercent, DocumentSettings Document, ImportSettings Import, int Listed, int Recent);

/// <summary>
/// What the app was set to, for an error report, in kinds and counts only: whether there's a source and what kind it
/// is, never its path or name, and how many items are ticked, never which.
/// </summary>
internal static class StateSummary
{
  /// <summary>The report's lines on what the app was set to.</summary>
  public static IReadOnlyList<string> Describe(AppState state)
  {
    ArgumentNullException.ThrowIfNull(state);
    var document = state.Document;
    var import = state.Import;
    var choose = import.ChooseFrom == ChooseFrom.Manifest
      ? $"a manifest ({KindOf(import.Manifest)})"
      : string.Create(CultureInfo.InvariantCulture, $"the list, {state.Listed} listed and {import.Items.Count} ticked");
    return
    [
      $"Tab showing: {state.Tab}",
      $"A run going: {YesNo(state.Running)}",
      string.Create(CultureInfo.InvariantCulture, $"Screen scale: {state.ScalePercent}%"),
      $"Document tab: input {KindOf(document.Input)}, subfolders {YesNo(document.Recursive)}; " +
        $"vbadoc {Ticked(document.Vbadoc.Run)}; pqdoc {Ticked(document.Pqdoc.Run)}, " +
        string.Create(CultureInfo.InvariantCulture, $"{document.Pqdoc.Formats.Count} formats"),
      $"Import tab: source {KindOf(import.Source)}, subfolders {YesNo(import.Recursive)}; {import.Tool.Name()}; " +
        $"choosing from {choose}; writing to {Destination(import.WriteTo)}",
      string.Create(CultureInfo.InvariantCulture, $"Recent workbooks and folders: {state.Recent}"),
    ];
  }

  /// <summary>What a path, as typed, is: "none", "a folder", "a file" or "not there". Never the path itself.</summary>
  public static string KindOf(string typed)
  {
    ArgumentNullException.ThrowIfNull(typed);
    var path = DocumentCommands.Tidy(typed);
    return path.Length == 0 ? "none"
      : Directory.Exists(path) ? "a folder"
      : File.Exists(path) ? "a file"
      : "not there";
  }

  private static string Destination(ImportDestination destination) => destination switch
  {
    ImportDestination.NewWorkbook => "a new workbook",
    ImportDestination.Copy => "a copy of a target",
    ImportDestination.InPlace => "the target itself",
    _ => "a bundle folder",
  };

  private static string YesNo(bool yes) => yes ? "yes" : "no";

  private static string Ticked(bool ticked) => ticked ? "ticked" : "not ticked";
}
