using System.CommandLine;
using CodeDocumenter.Shared.CommandLine;

namespace CodeDocumenter.Shared.Tests.CommandLine;

public sealed class CommandHelpTests
{
  private static readonly string NewLine = Environment.NewLine;

  /// <summary>A command shaped like the tools': an input, and an import with a source of its own.</summary>
  private static RootCommand Tool()
  {
    var root = new RootCommand("Documents a workbook.")
    {
      new Argument<string>("input") { Description = "The workbook." },
      new Option<string?>("--output", "-o"),
      new Command("import", "Takes items out of a source.")
      {
        new Argument<string>("source") { Description = "Where they come from." },
        new Option<string?>("--items", "-i"),
      },
    };
    root.SetAction(_ => { });
    CommandHelp.Apply(root, "tool");
    return root;
  }

  private static string Help(RootCommand root, params string[] args)
  {
    var output = new StringWriter();
    var result = root.Parse(args);

    Assert.IsType<CommandHelp.ToolHelpAction>(result.Action);
    Assert.Equal(0, result.Invoke(new InvocationConfiguration { Output = output, Error = new StringWriter() }));
    return output.ToString();
  }

  [Fact]
  public void The_tools_help_shows_it_run_with_its_input_and_each_subcommand_run_after_its_name()
  {
    var help = Help(Tool(), "--help");

    Assert.Contains($"Usage:{NewLine}  tool <input> [options]{NewLine}  tool import <source> [options]{NewLine}{NewLine}", help, StringComparison.Ordinal);
    Assert.Contains("  <input>  The workbook.", help, StringComparison.Ordinal);
    Assert.Contains("import <source>", help, StringComparison.Ordinal);
    Assert.DoesNotContain("[command]", help, StringComparison.Ordinal);
  }

  [Fact]
  public void A_subcommands_help_shows_it_run_after_the_tools_name_with_its_own_arguments_only()
  {
    var root = Tool();

    var help = Help(root, "import", "-h");

    Assert.Contains($"Usage:{NewLine}  tool import <source> [options]{NewLine}{NewLine}", help, StringComparison.Ordinal);
    Assert.Contains("  <source>  Where they come from.", help, StringComparison.Ordinal);
    Assert.DoesNotContain("<input>", help, StringComparison.Ordinal);
    Assert.False(root.Arguments.Single().Hidden);
  }

  [Fact]
  public void Anything_before_a_subcommand_is_refused_saying_how_its_run()
  {
    var errors = Tool().Parse(["Sales.xlsm", "import", "Other.xlsm", "-i", "A"]).Errors;

    Assert.Equal(
      "Nothing goes between tool and import, so 'Sales.xlsm' can't be there. Run it as tool import <source> [options].",
      Assert.Single(errors).Message);
  }

  [Fact]
  public void A_subcommand_run_after_the_tools_name_is_fine()
  {
    Assert.Empty(Tool().Parse(["import", "Other.xlsm", "-i", "A"]).Errors);
  }

  [Fact]
  public void The_brief_usage_after_a_mistake_is_the_commands_own()
  {
    var root = Tool();

    Assert.Equal($"Usage: tool <input> [options]{NewLine}Run 'tool --help' for the full list of options.", CommandHelp.Brief(root, "tool"));
    Assert.Equal(
      $"Usage: tool import <source> [options]{NewLine}Run 'tool import --help' for the full list of options.",
      CommandHelp.Brief(root.Subcommands.Single(), "tool"));
  }
}
