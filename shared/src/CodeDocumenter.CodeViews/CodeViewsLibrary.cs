using System.Reflection;

namespace CodeDocumenter.CodeViews;

/// <summary>
/// The library of views that tools documenting code share. It holds nothing yet but its place: step M2 of the
/// generalisation plan moves vbadoc's graph, class diagram and Calls tab here, and M3 the code model they draw.
/// </summary>
public static class CodeViewsLibrary
{
  /// <summary>The assembly that will embed the views' scripts and styles, for a page to inline them from.</summary>
  public static Assembly Assembly => typeof(CodeViewsLibrary).Assembly;
}
