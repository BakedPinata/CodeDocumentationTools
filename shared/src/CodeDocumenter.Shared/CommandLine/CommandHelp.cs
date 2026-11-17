using System.CommandLine;
using System.CommandLine.Help;
using System.CommandLine.Invocation;
using System.CommandLine.Parsing;

namespace CodeDocumenter.Shared.CommandLine;

/// <summary>
/// Help that shows each command as it's run: the tool's own command with its input, and each subcommand, such as
/// <c>import</c>, straight after the tool's name with its own arguments. System.CommandLine's own help puts the tool's
/// arguments before a subcommand, as in <c>vbadoc &lt;input&gt; import &lt;source&gt;</c>, and lists them as the
/// subcommand's, which isn't how it's run. Subcommands are one level deep, as the tools have them.
/// </summary>
public static class CommandHelp
{
  /// <summary>
  /// Makes <c>--help</c> show each command as it's run, and makes each subcommand refuse anything given before it, which
  /// it would otherwise take as the tool's own input and say nothing about.
  /// </summary>
  /// <param name="root">The tool's command, with its subcommands added.</param>
  /// <param name="tool">The tool's name, as it's typed.</param>
  public static void Apply(RootCommand root, string tool)
  {
    ArgumentNullException.ThrowIfNull(root);
    var help = root.Options.OfType<HelpOption>().Single();
    help.Action = new ToolHelpAction(tool, help.Action as HelpAction ?? new HelpAction());
    foreach (var command in root.Subcommands)
    {
      command.Validators.Add(result => RefuseWhatComesBefore(result, root, command, tool));
    }
  }

  /// <summary>How a command is run, a line for each way: the tool's own command, then each of its subcommands.</summary>
  /// <param name="command">The tool's command, or one of its subcommands.</param>
  /// <param name="tool">The tool's name, as it's typed.</param>
  public static IReadOnlyList<string> Usage(Command command, string tool)
  {
    ArgumentNullException.ThrowIfNull(command);
    return command is RootCommand
      ? [Line(command, tool), .. command.Subcommands.Where(c => !c.Hidden).Select(c => Line(c, tool))]
      : [Line(command, tool)];
  }

  /// <summary>What's shown after a mistake on the command line: how the command is run, and how to see the rest.</summary>
  /// <param name="command">The command the mistake was in.</param>
  /// <param name="tool">The tool's name, as it's typed.</param>
  public static string Brief(Command command, string tool)
  {
    ArgumentNullException.ThrowIfNull(command);
    return $"Usage: {Line(command, tool)}{Environment.NewLine}Run '{Name(command, tool)} --help' for the full list of options.";
  }

  private static string Name(Command command, string tool) => command is RootCommand ? tool : $"{tool} {command.Name}";

  private static string Line(Command command, string tool) =>
    string.Join(' ', [Name(command, tool), .. command.Arguments.Where(a => !a.Hidden).Select(Label), "[options]"]);

  // As System.CommandLine writes an argument: <name>, in brackets if it can be left out, and with ... if it takes more.
  private static string Label(Argument argument)
  {
    var label = $"<{argument.HelpName ?? argument.Name}>{(argument.Arity.MaximumNumberOfValues > 1 ? "..." : "")}";
    return argument.Arity.MinimumNumberOfValues == 0 ? $"[{label}]" : label;
  }

  private static void RefuseWhatComesBefore(CommandResult result, RootCommand root, Command command, string tool)
  {
    var before = root.Arguments.SelectMany(a => result.GetResult(a)?.Tokens ?? []).Select(t => $"'{t.Value}'").ToList();
    if (before.Count > 0)
    {
      result.AddError($"Nothing goes between {tool} and {command.Name}, so {string.Join(" and ", before)} can't be there. Run it as {Line(command, tool)}.");
    }
  }

  /// <summary>
  /// <c>--help</c>: System.CommandLine's own, with the usage as <see cref="Usage"/> says, and a subcommand's arguments
  /// without the tool's. Its own help can't be laid out otherwise, as it keeps how to from other code.
  /// </summary>
  /// <param name="tool">The tool's name, as it's typed.</param>
  /// <param name="help">System.CommandLine's own.</param>
  public sealed class ToolHelpAction(string tool, HelpAction help) : SynchronousCommandLineAction
  {
    /// <inheritdoc />
    public override bool ClearsParseErrors => true;

    /// <inheritdoc />
    public override int Invoke(ParseResult parseResult)
    {
      ArgumentNullException.ThrowIfNull(parseResult);
      var command = parseResult.CommandResult.Command;
      var configuration = parseResult.InvocationConfiguration;
      var output = configuration.Output;

      // The tool's arguments are hidden while a subcommand's help is written, so they aren't listed as its own.
      var hidden = command is RootCommand ? [] : parseResult.RootCommandResult.Command.Arguments.Where(a => !a.Hidden).ToList();
      var written = new StringWriter();
      try
      {
        hidden.ForEach(a => a.Hidden = true);
        configuration.Output = written;
        help.Invoke(parseResult);
      }
      finally
      {
        configuration.Output = output;
        hidden.ForEach(a => a.Hidden = false);
      }

      output.Write(WithUsage(written.ToString(), Usage(command, tool)));
      return 0;
    }

    // The lines under "Usage:", up to the blank line that ends the section, are put in place of its own.
    private static string WithUsage(string help, IReadOnlyList<string> usage)
    {
      var lines = help.Split(Environment.NewLine).ToList();
      var heading = lines.IndexOf("Usage:");
      if (heading < 0)
      {
        return help;
      }

      var end = lines.FindIndex(heading + 1, line => line.Length == 0);
      lines.RemoveRange(heading + 1, (end < 0 ? lines.Count : end) - heading - 1);
      lines.InsertRange(heading + 1, usage.Select(line => $"  {line}"));
      return string.Join(Environment.NewLine, lines);
    }
  }
}
