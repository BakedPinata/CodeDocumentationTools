using CodeDocumenter.Shared.Pages;

namespace CodeDocumenter.CodeViews;

/// <summary>
/// What a tool puts into its <see cref="ShellPage"/> to get a page of code with the views this library holds: the
/// dependency graph, the class diagram and the Calls tab, over the page script they talk to. The tool brings its
/// data, its search box, its icon, how its source is read and highlighted, and anything more it wants in the page.
/// </summary>
public static class CodeViewsPage
{
  private static readonly System.Reflection.Assembly Resources = typeof(CodeViewsPage).Assembly;

  /// <summary>The views' tabs, in order.</summary>
  public static IReadOnlyList<ShellTab> Tabs { get; } =
    [new ShellTab("graph", "Dependency graph"), new ShellTab("classes", "Class diagram"), new ShellTab("calls", "Calls")];

  /// <summary>The views, as HTML, for <see cref="ShellPage.ViewsHtml"/>.</summary>
  public static string ViewsHtml => EmbeddedResources.ReadText(Resources, "Views/views.html");

  /// <summary>The views' styles and the source panel's, for <see cref="ShellPage.Stylesheet"/>. A tool's own rules go after.</summary>
  public static string Stylesheet => EmbeddedResources.ReadText(Resources, "Views/views.css");

  /// <summary>
  /// The page's script. It reads what the tool says of itself in <c>window.codeViewsPage</c>, so the tool's script
  /// that sets that, and anything it names (a highlighter), comes before this one.
  /// </summary>
  public static PageScript PageScript { get; } = new(Resources, "Views/code-page.js");

  /// <summary>The views' scripts, in load order, after <see cref="PageScript"/>.</summary>
  public static IReadOnlyList<PageScript> ViewScripts { get; } =
    [new(Resources, "Views/graph.js"), new(Resources, "Views/classes.js"), new(Resources, "Views/calls.js")];

  /// <summary>
  /// Every script of a page of code, in load order: the tool's own (what sets <c>window.codeViewsPage</c>, and what
  /// that names), then the page's, then the views'.
  /// </summary>
  public static IReadOnlyList<PageScript> Scripts(IEnumerable<PageScript> toolScripts)
  {
    ArgumentNullException.ThrowIfNull(toolScripts);
    return [.. toolScripts, PageScript, .. ViewScripts];
  }
}
