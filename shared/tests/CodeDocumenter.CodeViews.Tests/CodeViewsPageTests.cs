using CodeDocumenter.Shared.Pages;

namespace CodeDocumenter.CodeViews.Tests;

/// <summary>What a tool takes from this library to make a page of code (step M2 of the generalisation plan).</summary>
public sealed class CodeViewsPageTests
{
  [Fact]
  public void Every_tab_has_its_view_in_the_views_html_and_the_first_is_the_one_shown()
  {
    var html = CodeViewsPage.ViewsHtml;

    Assert.Equal(["graph", "classes", "calls"], CodeViewsPage.Tabs.Select(t => t.Name));
    Assert.All(CodeViewsPage.Tabs, tab => Assert.Contains($"id=\"view-{tab.Name}\" role=\"tabpanel\" aria-labelledby=\"tab-{tab.Name}\"", html, StringComparison.Ordinal));
    Assert.DoesNotContain("id=\"view-graph\" role=\"tabpanel\" aria-labelledby=\"tab-graph\" hidden", html, StringComparison.Ordinal);
    Assert.Contains("id=\"view-classes\" role=\"tabpanel\" aria-labelledby=\"tab-classes\" hidden", html, StringComparison.Ordinal);
  }

  [Fact]
  public void The_tools_scripts_come_first_then_the_pages_then_the_views()
  {
    var tool = new PageScript(typeof(CodeViewsPageTests).Assembly, "Page/page.js");

    var scripts = CodeViewsPage.Scripts([tool]);

    Assert.Equal(["Page/page.js", "Views/code-page.js", "Views/graph.js", "Views/classes.js", "Views/calls.js"], scripts.Select(s => s.Name));
    Assert.Same(tool, scripts[0]);
  }

  [Fact]
  public void Every_script_is_embedded_and_can_be_inlined_in_a_page()
  {
    var scripts = PageBuilder.Scripts([CodeViewsPage.PageScript, .. CodeViewsPage.ViewScripts]);

    Assert.Contains("// code-page.js:", scripts, StringComparison.Ordinal);
    Assert.Contains("// graph.js:", scripts, StringComparison.Ordinal);
    Assert.Contains("// classes.js:", scripts, StringComparison.Ordinal);
    Assert.Contains("// calls.js:", scripts, StringComparison.Ordinal);
  }

  [Fact]
  public void The_stylesheet_has_the_views_rules()
  {
    var css = CodeViewsPage.Stylesheet;

    Assert.Contains(".graph-svg", css, StringComparison.Ordinal);
    Assert.Contains(".classes-svg", css, StringComparison.Ordinal);
    Assert.Contains(".calls-tree", css, StringComparison.Ordinal);
  }
}
