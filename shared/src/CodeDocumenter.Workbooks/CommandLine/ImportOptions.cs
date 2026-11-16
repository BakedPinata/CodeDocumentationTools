using CodeDocumenter.Shared.CommandLine;
using System.CommandLine;
using System.CommandLine.Parsing;
using CodeDocumenter.Shared.Files;

namespace CodeDocumenter.Workbooks.CommandLine;

/// <summary>What an <c>import</c> command was asked to do.</summary>
/// <param name="Source">The workbook (or pqdoc's folder) to read, as a full path.</param>
/// <param name="Manifest">The manifest to read the choice from, as a full path, or null.</param>
/// <param name="Items">The items named on the command line.</param>
/// <param name="Bundle">The folder to write a bundle into, as a full path, or null to only say what would come.</param>
/// <param name="OnExists">What to do when the folder holds an earlier bundle.</param>
/// <param name="Recursive">Whether a folder's subfolders are read too (<c>--recursive</c>).</param>
public sealed record ImportRequest(
  string Source, string? Manifest, IReadOnlyList<string> Items, string? Bundle, OnExistsBehaviour OnExists, bool Recursive = false);

/// <summary>
/// The <c>import</c> command, the same in every tool: which items, from a manifest the page's Export tab wrote or by name,
/// and where they go.
/// </summary>
public sealed class ImportOptions
{
  /// <summary>Creates the command and its options.</summary>
  /// <param name="source">What the source is, for the help text.</param>
  /// <param name="items">What the items are, for the help text, such as "modules and classes".</param>
  /// <param name="files">
  /// The files read from a folder, for <c>--recursive</c>'s help text, such as ".m files"; null for a tool that can't take
  /// from a folder, which then has no <c>--recursive</c>.
  /// </param>
  public ImportOptions(string source, string items, string? files = null)
  {
    Source = new Argument<string>("source") { Description = source };

    Manifest = new Option<string?>("--manifest", "-m")
    {
      Description = $"The {items} to take: a manifest downloaded from the page's Export tab.",
      HelpName = "file",
    };

    Items = new Option<string?>("--items", "-i")
    {
      Description = $"The {items} to take, by name, separated by commas, in place of --manifest.",
      HelpName = "names",
    };

    Bundle = new Option<string?>("--bundle", "-b")
    {
      Description = "A folder to write them into, with a README on importing them by hand. Without it, only says what would come.",
      HelpName = "folder",
    };

    Overwrite = new OverwriteOptions("an earlier bundle", "an earlier bundle");
    Recursive = files is null ? null : FolderOptions.CreateRecursive(files);

    Command = new Command("import", $"Takes chosen {items}, with what they use, out of a source. Run it with a manifest from the page's Export tab.")
    {
      Source,
      Manifest,
      Items,
      Bundle,
      Overwrite.OnExists,
      Overwrite.Force,
    };
    if (Recursive is not null)
    {
      Command.Options.Add(Recursive);
    }

    Command.Validators.Add(Validate);
  }

  /// <summary>The <c>import</c> command.</summary>
  public Command Command { get; }

  /// <summary>The source.</summary>
  public Argument<string> Source { get; }

  /// <summary><c>--manifest</c> / <c>-m</c>.</summary>
  public Option<string?> Manifest { get; }

  /// <summary><c>--items</c> / <c>-i</c>.</summary>
  public Option<string?> Items { get; }

  /// <summary><c>--bundle</c> / <c>-b</c>.</summary>
  public Option<string?> Bundle { get; }

  /// <summary><c>--recursive</c> / <c>-r</c>, or null for a tool that can't take from a folder.</summary>
  public Option<bool>? Recursive { get; }

  /// <summary><c>--on-exists</c> and <c>--force</c> / <c>-y</c>.</summary>
  public OverwriteOptions Overwrite { get; }

  /// <summary>What was asked for, with its paths made full.</summary>
  public ImportRequest Bind(ParseResult parseResult)
  {
    ArgumentNullException.ThrowIfNull(parseResult);
    var manifest = parseResult.GetValue(Manifest);
    var bundle = parseResult.GetValue(Bundle);
    var items = parseResult.GetValue(Items)?.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries) ?? [];
    return new ImportRequest(
      Path.GetFullPath(parseResult.GetValue(Source)!),
      manifest is null ? null : Path.GetFullPath(manifest),
      items,
      bundle is null ? null : Path.GetFullPath(bundle),
      Overwrite.Bind(parseResult),
      Recursive is not null && parseResult.GetValue(Recursive));
  }

  private void Validate(CommandResult result)
  {
    var manifest = result.GetValue(Manifest);
    var items = result.GetValue(Items);
    if (manifest is null == items is null)
    {
      result.AddError("Say what to take with one of --manifest (-m) or --items (-i).");
    }

    Overwrite.Validate(result);
  }
}
