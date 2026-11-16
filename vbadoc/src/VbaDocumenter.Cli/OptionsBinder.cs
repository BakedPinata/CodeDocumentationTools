using System.CommandLine;
using VbaDocumenter.Core.Output;
using VbaDocumenter.Cli.Logging;

namespace VbaDocumenter.Cli;

/// <summary>The outcome of turning a parse result into <see cref="CommandLineOptions"/>.</summary>
/// <param name="Options">The options, or null when there are errors.</param>
/// <param name="Errors">What was wrong with the arguments (exit code 2).</param>
sealed record BindingResult(CommandLineOptions? Options, IReadOnlyList<string> Errors);

/// <summary>Turns a successful parse into <see cref="CommandLineOptions"/>, resolving paths.</summary>
static class OptionsBinder
{
  public static BindingResult Bind(ParseResult parseResult, CliDefinition cli)
  {
    var errors = new List<string>();

    // A folder of module files is documented where it is: the page goes in it, and nothing else is written.
    var workbook = ResolvePath(parseResult.GetValue(cli.Workbook), "input", errors);
    var isFolder = workbook is not null && Directory.Exists(workbook);
    var output = parseResult.GetValue(cli.Output);
    var outputFolder = output is not null
        ? ResolvePath(output, "--output", errors)
        : workbook is null ? null : isFolder ? workbook : Path.Combine(Path.GetDirectoryName(workbook)!, "VBAProject");

    if (workbook is null || outputFolder is null) return new BindingResult(null, errors);

    var outputName = parseResult.GetValue(cli.OutputName) ?? (isFolder ? HtmlPageWriter.DefaultFolderName(workbook) : HtmlPageWriter.DefaultName(workbook));
    var options = new CommandLineOptions(
        workbook, outputFolder, outputName, parseResult.GetValue(cli.EmbedSource), cli.Overwrite.Bind(parseResult), GetVerbosity(parseResult, cli),
        parseResult.GetValue(cli.Folder.Recursive), parseResult.GetValue(cli.Folder.FolderGroups));
    return new BindingResult(options, errors);
  }

  /// <summary>
  /// Reads the requested verbosity. Safe to call on a parse result with errors, since it is needed to set up logging
  /// before they are reported. Conflicting flags fall back to normal.
  /// </summary>
  public static LogVerbosity GetVerbosity(ParseResult parseResult, CliDefinition cli) =>
      (parseResult.GetValue(cli.Verbose), parseResult.GetValue(cli.Quiet)) switch
      {
        (true, false) => LogVerbosity.Verbose,
        (false, true) => LogVerbosity.Quiet,
        _ => LogVerbosity.Normal,
      };

  static string? ResolvePath(string? value, string name, List<string> errors)
  {
    if (string.IsNullOrWhiteSpace(value))
    {
      errors.Add($"{name} must not be empty.");
      return null;
    }

    try
    {
      return Path.GetFullPath(value);
    }
    catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
    {
      errors.Add($"{name} '{value}' is not a valid path: {ex.Message}");
      return null;
    }
  }
}
