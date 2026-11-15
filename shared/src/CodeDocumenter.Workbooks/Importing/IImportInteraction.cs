using CodeDocumenter.Shared.Files;

namespace CodeDocumenter.Workbooks.Importing;

/// <summary>What <c>import</c> asks the person running the tool, on top of what every tool asks.</summary>
public interface IImportInteraction : IUserInteraction
{
  /// <summary>Asks what to do with an item the target already has. Overwrite is only offered if it can be replaced.</summary>
  ClashAnswer AskClash(ImportClash clash);
}
