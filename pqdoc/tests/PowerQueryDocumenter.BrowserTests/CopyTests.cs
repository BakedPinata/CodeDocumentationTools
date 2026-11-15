using Microsoft.Playwright;

namespace PowerQueryDocumenter.BrowserTests;

/// <summary>
/// Copy as text (FR-HTML-26, step L6 of the generalisation plan), which is the shell's, as vbadoc's is: here, that
/// the queries the graph shows are copied as Markdown, with each query's description, what it uses and its M code
/// when asked.
/// </summary>
[Collection("Browser")]
public sealed class CopyTests(PageFixture fixture)
{
  private const string ClipboardKept = """
      Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: text => { window.copiedText = text; return Promise.resolve(); } } });
      """;

  private static Task<string> Text(IPage page, bool source = false) =>
    page.EvaluateAsync<string>("source => window.pqdoc.viewText({ source })", source);

  [Fact]
  public async Task FR_HTML_26_copy_puts_the_queries_the_graph_shows_on_the_clipboard()
  {
    await using var opened = await OpenedPage.OpenAsync(fixture, fixture.EmbeddedPage, initScript: ClipboardKept);
    var page = opened.Page;

    await page.ClickAsync("#copy > summary");
    await page.ClickAsync("#copy-view");
    await page.WaitForFunctionAsync("() => !document.getElementById('copy-view').disabled");

    var copied = await page.EvaluateAsync<string>("() => window.copiedText");
    Assert.StartsWith("# sales-model\n\n", copied);
    Assert.Contains("\nMade by pqdoc ", copied);
    Assert.Contains("\n\n## The page as it is now\n\n", copied);
    Assert.Contains("\n\n### 11 items\n\n", copied);
    Assert.Contains("\n\n#### Orders (Query)\n\nIntermediate query.\n\n- Uses: fnCleanText, Sales Region Map, Source\n", copied);
    Assert.Contains("\n\n#### fnCleanText (Function)\n\n", copied);
    Assert.DoesNotContain("```", copied); // no source unless asked for
    Assert.Equal($"Copied {copied.Length:N0} characters.", await page.TextContentAsync("#copy-status"));
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task FR_HTML_26_a_querys_M_code_is_included_when_asked_for()
  {
    await using var opened = await OpenedPage.OpenAsync(fixture, fixture.EmbeddedPage);

    var text = await Text(opened.Page, source: true);

    var orders = text[text.IndexOf("#### Orders (Query)", StringComparison.Ordinal)..];
    orders = orders[..orders.IndexOf("\n#### ", StringComparison.Ordinal)];
    Assert.Contains("\n\n```powerquery\n", orders);
    Assert.Contains("Table.NestedJoin", orders);
    Assert.EndsWith("\n```\n", orders);
    Assert.Empty(opened.Errors);
  }
}
