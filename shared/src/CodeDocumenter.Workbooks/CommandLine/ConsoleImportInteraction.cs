using CodeDocumenter.Shared.CommandLine;
using CodeDocumenter.Workbooks.Importing;

namespace CodeDocumenter.Workbooks.CommandLine;

/// <summary>The console's prompts, with the one <c>import</c> adds: what to do with an item the target already has.</summary>
public sealed class ConsoleImportInteraction : ConsoleUserInteraction, IImportInteraction
{
  /// <inheritdoc />
  public ClashAnswer AskClash(ImportClash clash)
  {
    ArgumentNullException.ThrowIfNull(clash);
    Console.Error.WriteLine(clash.Problem);
    if (clash.IfOverwritten is not null)
    {
      Console.Error.WriteLine($"  Overwrite: {clash.IfOverwritten}");
    }

    Console.Error.WriteLine($"  Skip: {clash.IfSkipped}");
    while (true)
    {
      Console.Error.Write(clash.CanOverwrite
        ? "[s]kip / [o]verwrite / [sa] skip all / [oa] overwrite all / [q]uit: "
        : "[s]kip / [sa] skip all / [q]uit: ");
      var answer = Console.ReadLine();

      // End of input (for example Ctrl+Z / Ctrl+D) is treated as quit.
      if (answer is null)
      {
        Console.Error.WriteLine();
        return ClashAnswer.Quit;
      }

      if (ReadClashAnswer(answer, clash.CanOverwrite) is { } read)
      {
        return read;
      }

      Console.Error.WriteLine(clash.CanOverwrite ? "Please answer s, o, sa, oa or q." : "Please answer s, sa or q.");
    }
  }

  /// <summary>What a typed answer to the clash prompt means, or null if it means nothing.</summary>
  public static ClashAnswer? ReadClashAnswer(string answer, bool canOverwrite)
  {
    ArgumentNullException.ThrowIfNull(answer);
    return string.Join(' ', answer.Split(' ', StringSplitOptions.RemoveEmptyEntries)).ToLowerInvariant() switch
    {
      "s" or "skip" => ClashAnswer.Skip,
      "sa" or "skip all" => ClashAnswer.SkipAll,
      "o" or "overwrite" when canOverwrite => ClashAnswer.Overwrite,
      "oa" or "overwrite all" when canOverwrite => ClashAnswer.OverwriteAll,
      "q" or "quit" => ClashAnswer.Quit,
      _ => null,
    };
  }
}
