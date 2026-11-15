using CodeDocumenter.Shared.Files;

namespace CodeDocumenter.Shared.CommandLine;

/// <summary>
/// Prompts on the console. Prompts go to stderr so that stdout stays clean. A tool with more to ask adds its
/// questions in a class of its own, as the workbook tools' import does.
/// </summary>
public class ConsoleUserInteraction : IUserInteraction
{
  /// <inheritdoc />
  public bool IsInteractive => !Console.IsInputRedirected && !Console.IsOutputRedirected && !Console.IsErrorRedirected;

  /// <inheritdoc />
  public OverwriteAnswer AskOverwrite(string path)
  {
    while (true)
    {
      Console.Error.Write($"'{path}' already exists. Overwrite? [y]es / [n]o / [a]ll / [q]uit: ");
      var answer = Console.ReadLine();

      // End of input (for example Ctrl+Z / Ctrl+D) is treated as quit.
      if (answer is null)
      {
        Console.Error.WriteLine();
        return OverwriteAnswer.Quit;
      }

      switch (answer.Trim().ToLowerInvariant())
      {
        case "y" or "yes":
          return OverwriteAnswer.Yes;
        case "n" or "no":
          return OverwriteAnswer.No;
        case "a" or "all":
          return OverwriteAnswer.All;
        case "q" or "quit":
          return OverwriteAnswer.Quit;
        default:
          Console.Error.WriteLine("Please answer y, n, a or q.");
          break;
      }
    }
  }

  /// <inheritdoc />
  public bool Confirm(string question)
  {
    while (true)
    {
      Console.Error.Write($"{question} [y]es / [n]o: ");
      var answer = Console.ReadLine();

      // End of input is treated as no: this is for what can't be undone.
      if (answer is null)
      {
        Console.Error.WriteLine();
        return false;
      }

      switch (answer.Trim().ToLowerInvariant())
      {
        case "y" or "yes":
          return true;
        case "n" or "no":
          return false;
        default:
          Console.Error.WriteLine("Please answer y or n.");
          break;
      }
    }
  }
}
