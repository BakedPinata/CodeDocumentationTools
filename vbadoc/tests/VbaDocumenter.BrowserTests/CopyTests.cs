using Microsoft.Playwright;

namespace VbaDocumenter.BrowserTests;

/// <summary>
/// Copy as text (step L6 of the generalisation plan): the Copy menu puts what the page shows on the clipboard as
/// Markdown, to paste into a chat: the view's items with their signatures and descriptions, the Calls tree as it's
/// opened, and the source if asked. It warns when the text is long, and shows it in a box when the browser won't copy.
/// </summary>
[Collection("Browser")]
public sealed class CopyTests(PageFixture fixture)
{
  /// <summary>A clipboard that keeps what's written to it where the test can read it.</summary>
  const string ClipboardKept = """
      Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: text => { window.copiedText = text; return Promise.resolve(); } } });
      """;

  /// <summary>A browser that refuses both ways of copying.</summary>
  const string ClipboardRefused = """
      Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: () => Promise.reject(new Error('refused')) } });
      document.execCommand = () => false;
      """;

  static Task<string> Text(IPage page, bool source = false) =>
      page.EvaluateAsync<string>("source => window.vbadoc.viewText({ source })", source);

  static async Task Copy(IPage page)
  {
    await page.ClickAsync("#copy > summary");
    await page.ClickAsync("#copy-view");
    await page.WaitForFunctionAsync("() => !document.getElementById('copy-view').disabled");
  }

  [Fact]
  public async Task Copy_puts_the_view_on_the_clipboard_with_each_item_and_says_how_much()
  {
    await using var opened = await OpenedPage.OpenAsync(fixture, fixture.ConditionsEmbedded, initScript: ClipboardKept);
    var page = opened.Page;
    await opened.SelectAsync("Described");

    await Copy(page);

    var copied = await page.EvaluateAsync<string>("() => window.copiedText");
    Assert.StartsWith("# Conditions.xlsm\n\n14 modules, ", copied);
    Assert.Contains("\nMade by vbadoc ", copied);
    Assert.Contains("\n\n## The page as it is now\n\nDependency graph · selected: Described (focused)\n\n### 1 item\n\n", copied);
    // The module, with its file as a note, not as code, and its description as written.
    Assert.EndsWith("#### Described (Standard module)\n\nDescribed.bas · 15 lines\n\nWorks out prices for the <shop>,\ntax included.\n", copied);
    Assert.Equal($"Copied {copied.Length:N0} characters.", await page.TextContentAsync("#copy-status"));
    Assert.True(await page.IsHiddenAsync("#copy-fallback"));
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task The_Calls_tab_copies_the_tree_as_it_is_opened_then_the_selection()
  {
    await using var opened = await OpenedPage.OpenAsync(fixture, fixture.ConditionsEmbedded);
    var page = opened.Page;
    await page.ClickAsync("#tab-calls");
    await page.ClickAsync("#calls-tree li.group > ul > li[data-id='Conditions.CheckBoth'] > .tree-row .tree-twisty");
    await page.ClickAsync("#calls-tree li[data-id='Conditions.IsB'] .calls-name");

    var text = await Text(page);

    Assert.Contains("\n\nCalls · selected: Conditions.IsB\n\n### Calls, from the entry points\n\n", text);
    // Groups, a procedure that isn't opened, and one that is: its blocks in the code's order, with the calls their own
    // statements make marked, and each call's line.
    Assert.Contains(string.Join("\n",
        "- Macros (4)",
        "  - AppInit.Startup (Sub, macro) …",
        "  - Conditions.CheckBoth (Sub, macro)",
        "    - `If IsA(1) Or IsB(2) Then`, line 6",
        "      - condition: Conditions.IsA (Function), line 6",
        "      - condition: Conditions.IsB (Function), line 6",
        "    - `Else`, line 8",
        "      - Conditions.DoB (Sub), line 9",
        "    - `Do While IsA(x)`, line 11",
        "      - condition: Conditions.IsA (Function), line 11",
        "  - Conditions.ManyCalls (Sub, macro) …"), text);
    Assert.Contains("\n- Event handlers (4)\n  - Order.Class_Initialize (Sub, event handler)\n", text);
    // Then the selection, with its signature as code.
    Assert.EndsWith("### 1 item\n\n#### Conditions.IsB (Function)\n\n```vba\nPrivate Function IsB(n As Long) As Boolean\n```\n\n- Used by: CheckBoth\n", text);
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task Called_by_copies_each_caller_with_the_block_its_call_is_in()
  {
    await using var opened = await OpenedPage.OpenAsync(fixture, fixture.SampleEmbedded);
    var page = opened.Page;
    await page.ClickAsync("#tab-calls");
    await page.ClickAsync("[data-calls-direction='calledBy']");
    await opened.SelectAsync("Logger.LogInfo");

    // The procedure the tree starts from is open to begin with.
    var open = await Text(page);
    await page.Locator("#calls-tree li[data-id='Logger.LogInfo'] > .tree-row .tree-twisty").First.ClickAsync();
    var closed = await Text(page);

    Assert.Contains("### Called by, into Logger.LogInfo\n\nWhat calls each procedure, as far as the tree is opened.", open);
    Assert.Contains("\n- Logger.LogInfo (Sub) …\n", closed);
    Assert.Contains("\n- Logger.LogInfo (Sub)\n  - AppInit.Startup (Sub), line 8 …\n", open);
    Assert.Contains("\n  - Main.ProcessOrders (Sub), line 26, in `Case \"HOLD\"` …\n", open);
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task Source_is_included_when_asked_for_as_a_code_block_of_the_items_own_lines()
  {
    await using var opened = await OpenedPage.OpenAsync(fixture, fixture.ConditionsEmbedded);
    var page = opened.Page;
    await page.ClickAsync("#tab-calls");
    await opened.SelectAsync("Described.WithTax");

    var without = await Text(page);
    var with = await Text(page, source: true);

    Assert.DoesNotContain("Round(net", without);
    Assert.EndsWith(string.Join("\n",
        "#### Described.WithTax (Function)",
        "",
        "```vba",
        "Public Function WithTax(net As Double) As Double",
        "```",
        "",
        "The price with tax, to two places.",
        "",
        "Lines 7 to 11:",
        "",
        "```vba",
        "' Adds the tax to a net price.",
        "Public Function WithTax(ByVal net As Double) As Double",
        "Attribute WithTax.VB_Description = \"The price with tax, to two places.\"",
        "    WithTax = Round(net * 1.1, 2)",
        "End Function",
        "```",
        ""), with);
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task Without_the_source_loaded_the_text_says_why_it_is_not_there()
  {
    await using var opened = await OpenedPage.OpenAsync(fixture, fixture.SampleFolder);
    await opened.SelectAsync("Logger");

    var text = await Text(opened.Page, source: true);

    Assert.EndsWith("\n\nThe source isn't included: it isn't loaded in the page. Use Load folder first.\n", text);
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task Names_and_code_that_mean_something_to_Markdown_are_copied_as_written()
  {
    await using var opened = await OpenedPage.OpenAsync(fixture, fixture.AwkwardEmbedded);
    await opened.SelectAsync(PageFixture.AwkwardModule);

    var text = await Text(opened.Page, source: true);

    Assert.Contains($"#### {PageFixture.AwkwardModule} (Standard module)", text);
    // The source holds a backtick, so nothing in it can end its block early: the fence is three, the most in it is one.
    Assert.Contains("\n```vba\nAttribute VB_Name", text);
    Assert.Contains("Public Const MARKDOWN As String = \"**x** __y__ [l](u) `c`\"\n", text);
    Assert.EndsWith("End Function\n```\n", text);
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task A_long_text_is_copied_with_a_warning_and_a_way_to_shorten_it()
  {
    await using var opened = await OpenedPage.OpenAsync(fixture, fixture.LargeEmbedded, initScript: ClipboardKept);
    var page = opened.Page;
    await page.ClickAsync("#copy > summary");
    await page.CheckAsync("#copy-source");

    await page.ClickAsync("#copy-view");
    await page.WaitForFunctionAsync("() => !document.getElementById('copy-view').disabled", null, new() { Timeout = 60_000 });

    var copied = await page.EvaluateAsync<string>("() => window.copiedText");
    Assert.True(copied.Length > 30_000);
    Assert.Contains("\n\n### 300 items\n\n", copied);
    Assert.Contains("… and 100 more, not listed. Tick fewer items to list them all.", copied);
    Assert.Equal($"Copied {copied.Length:N0} characters. That's long for a chat: tick fewer items in the explorer, or leave out the source.",
        await page.TextContentAsync("#copy-status"));
    Assert.True(await page.EvaluateAsync<bool>("() => document.getElementById('copy-status').classList.contains('warning')"));
    // The choice is remembered.
    Assert.Equal("true", await page.EvaluateAsync<string>("() => localStorage.getItem('vbadoc-copy-source')"));
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task When_the_browser_refuses_the_clipboard_the_text_is_shown_in_a_box_to_copy_by_hand()
  {
    await using var opened = await OpenedPage.OpenAsync(fixture, fixture.SampleEmbedded, initScript: ClipboardRefused);
    var page = opened.Page;
    await opened.SelectAsync("Logger");

    await Copy(page);

    Assert.True(await page.IsVisibleAsync("#copy-fallback"));
    var shown = await page.InputValueAsync("#copy-text");
    Assert.Equal(await Text(page), shown);
    Assert.Contains("#### Logger (Standard module)", shown);
    // It's selected, ready for Ctrl+C.
    Assert.True(await page.EvaluateAsync<bool>(
        "() => { const box = document.getElementById('copy-text'); return document.activeElement === box && box.selectionStart === 0 && box.selectionEnd === box.value.length; }"));
    Assert.Equal($"{shown.Length:N0} characters.", await page.TextContentAsync("#copy-status"));

    // Opening the menu again starts afresh.
    await page.ClickAsync("#copy > summary");
    await page.ClickAsync("#copy > summary");
    Assert.True(await page.IsHiddenAsync("#copy-fallback"));
    Assert.Empty(opened.Errors);
  }
}
