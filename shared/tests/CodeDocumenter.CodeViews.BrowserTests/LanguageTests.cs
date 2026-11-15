using Microsoft.Playwright;

namespace CodeDocumenter.CodeViews.BrowserTests;

/// <summary>
/// The views draw what a language describes (step M3 of the generalisation plan): here a made-up one, Toy, whose
/// units are gadgets of two kinds, whose members are parts, and which has a relationship of its own. Nothing the
/// page shows comes from VBA.
/// </summary>
[Collection("Browser")]
public sealed class LanguageTests(ToyPageFixture fixture)
{
  private static Task<string[]> Texts(IPage page, string selector) =>
    page.Locator(selector).EvaluateAllAsync<string[]>("els => els.map(el => el.textContent.trim())");

  [Fact]
  public async Task The_page_uses_the_languages_words_and_none_of_another_languages()
  {
    await using var opened = await fixture.OpenAsync(fixture.Embedded);
    var page = opened.Page;

    Assert.Equal("Select a gadget or part to see its source.", await page.InnerTextAsync("#side-panel-body .source-message"));
    Assert.Equal(["Gadgets", "Parts"], await Texts(page, "#view-graph [data-level]"));
    await page.ClickAsync("#explorer-show > summary");
    Assert.Equal(["Gadgets", "Parts", "References", "Lines", "Other", "Widely used"],
      (await Texts(page, "#explorer-show-menu h3, #explorer-show-menu legend")).Select(t => t.Replace(" All · None", "", StringComparison.Ordinal)));
    await page.Keyboard.PressAsync("Escape");

    // Every tab, with something selected and with nothing: no word of VBA's or a workbook's anywhere a person reads.
    foreach (var id in new[] { null, "Widget", "Widget.Stop" })
    {
      await opened.SelectAsync(id);
      foreach (var tab in new[] { "graph", "classes", "calls" })
      {
        await page.ClickAsync($"#tab-{tab}");
        var read = await page.EvaluateAsync<string>(@"() => {
          const parts = [document.body.innerText];
          document.querySelectorAll('body [title], body [aria-label], body [placeholder]').forEach(el =>
            parts.push(el.getAttribute('title') || '', el.getAttribute('aria-label') || '', el.getAttribute('placeholder') || ''));
          document.querySelectorAll('body svg title').forEach(el => parts.push(el.textContent));
          return parts.join('\n');
        }");
        Assert.DoesNotMatch(@"(?i)\bmodules?\b|\bVBA\b|vbadoc|workbook|macro|UserForm|late-bound|Application\.Run|\bunits?\b", read);
      }
    }

    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task The_explorer_lays_gadgets_out_by_the_languages_kinds_and_in_their_groups()
  {
    await using var opened = await fixture.OpenAsync(fixture.Embedded);
    var page = opened.Page;

    // Groups first, as the model has some: a shelf inside a shelf, and what's on none.
    Assert.Equal(["Shelves", "Types"], await Texts(page, "#explorer-layouts [role='radio']"));
    await page.ClickAsync("#explorer-expand-all");
    Assert.Equal(["Shop", "No group", "Widely used"], await Texts(page, "#explorer-tree > li > .tree-row .tree-label")); // the last is the shell's
    Assert.Equal(["Back", "Gizmo"], await Texts(page, "#explorer-tree > li:first-child > ul > li > .tree-row .tree-label"));
    Assert.Equal(["Widget"], await Texts(page, "#explorer-tree > li:first-child > ul > li:first-child > ul > li > .tree-row .tree-label"));
    Assert.Equal(["Blueprint"], await Texts(page, "#explorer-tree > li:nth-child(2) > ul > li > .tree-row .tree-label"));

    await page.ClickAsync("#explorer-layouts [role='radio']:has-text('Types')");
    Assert.Equal(["Things", "Plans", "Widely used"], await Texts(page, "#explorer-tree > li > .tree-row .tree-label"));
    Assert.Equal(["Gizmo", "Widget"], await Texts(page, "#explorer-tree > li:first-child > ul > li > .tree-row .tree-label"));
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task The_graph_draws_the_languages_kinds_and_its_own_relationship()
  {
    await using var opened = await fixture.OpenAsync(fixture.Embedded);
    var page = opened.Page;

    Assert.Equal(["Thing", "Plan", "Uses", "Extends"], await Texts(page, "#graph-legend-body .legend-item"));
    Assert.Equal(3, await page.Locator("#graph-svg g.node").CountAsync());
    // A kind's colour slot and whether it's hollow are the language's.
    Assert.Equal("node group-3 hollow", (await page.GetAttributeAsync("#graph-svg g.node[data-id='Blueprint']", "class"))!.Trim());
    Assert.StartsWith("node group-1", await page.GetAttributeAsync("#graph-svg g.node[data-id='Gizmo']", "class"));

    // Its own relationship is a line like any other: dashed, as it says, with its own arrowhead and words.
    var extends = page.Locator("#graph-svg path.edge.extends");
    Assert.Equal(1, await extends.CountAsync());
    Assert.Equal("6px, 4px", await extends.EvaluateAsync<string>("el => getComputedStyle(el).strokeDasharray"));
    Assert.Equal("url(#arrow-extends)", await extends.GetAttributeAsync("marker-end"));
    Assert.Equal("Gizmo extends Blueprint", await extends.Locator("title").TextContentAsync());
    Assert.Equal("none", await page.Locator("#graph-svg path.edge.uses").First.EvaluateAsync<string>("el => getComputedStyle(el).strokeDasharray"));

    // At member level, the legend's categories and colours are the language's too.
    await page.ClickAsync("#view-graph [data-level='members']");
    Assert.Equal(["Steps", "Facts", "The gadget itself", "Unreferenced", "Thing", "Plan", "Uses", "Extends"], await Texts(page, "#graph-legend-body .legend-item"));
    Assert.Equal(["Part", "Gadget (colour and outline)", "Edges"], await Texts(page, "#graph-legend-body h3"));
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task The_Show_menu_offers_the_languages_kinds_and_lines_and_they_take_effect()
  {
    await using var opened = await fixture.OpenAsync(fixture.Embedded);
    var page = opened.Page;
    await page.ClickAsync("#explorer-show > summary");

    Assert.Equal(["unit:thing", "unit:plan", "member:step", "member:fact", "private", "line:uses", "line:extends", "unconnected"],
      await page.Locator("#explorer-show-menu [data-choice]").EvaluateAllAsync<string[]>("els => els.map(el => el.getAttribute('data-choice'))"));

    await page.UncheckAsync("#explorer-show [data-choice='unit:plan']");
    await page.WaitForFunctionAsync("() => document.querySelectorAll('#graph-svg g.node').length === 2");
    Assert.Equal(0, await page.Locator("#graph-svg path.edge.extends").CountAsync());
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task The_class_diagram_takes_its_stereotypes_and_lines_from_the_language()
  {
    await using var opened = await fixture.OpenAsync(fixture.Embedded);
    var page = opened.Page;
    await page.ClickAsync("#tab-classes");
    await page.WaitForFunctionAsync("() => window.codeViewsClasses.drawn()");

    Assert.Equal("«thing»", await page.TextContentAsync("#classes-svg g.class-box[data-unit='Widget'] .stereotype"));
    Assert.Equal("«plan»", await page.TextContentAsync("#classes-svg g.class-box[data-unit='Blueprint'] .stereotype"));
    Assert.Equal(["+Spin()", "-Weight = 3", "+Stop()"],
      (await Texts(page, "#classes-svg g.class-box[data-unit='Widget'] .class-line text")).OrderBy(t => t == "-Weight = 3" ? 1 : t == "+Spin()" ? 0 : 2));
    Assert.Equal(["Thing", "Plan", "Extends", "Uses"], await Texts(page, "#classes-legend-body .legend-item"));
    Assert.Contains("+ open, - shut.", await page.InnerTextAsync("#classes-legend-body"));

    var inherit = page.Locator("#classes-svg path.class-edge.inherit");
    Assert.Equal(1, await inherit.CountAsync());
    Assert.Equal("url(#classes-arrow-inherit)", await inherit.GetAttributeAsync("marker-end"));
    Assert.Equal(1, await page.Locator("#classes-svg defs marker#classes-arrow-inherit").CountAsync());
    Assert.Equal("Gizmo extends Blueprint", await inherit.Locator("title").TextContentAsync());
    Assert.Equal("5px, 3px", await page.Locator("#classes-svg path.class-edge.use").First.EvaluateAsync<string>("el => getComputedStyle(el).strokeDasharray"));
    Assert.Equal("List each gadget's parts in its box", await page.GetAttributeAsync("#classes-members", "title"));
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task The_Calls_tab_starts_from_the_languages_entry_points_and_follows_its_relationship()
  {
    await using var opened = await fixture.OpenAsync(fixture.Embedded);
    var page = opened.Page;
    await page.ClickAsync("#tab-calls");

    Assert.Equal(["Starting points"], await Texts(page, "#calls-tree > li.group > .tree-row .tree-label"));
    Assert.Equal("Calls the code makes directly, in its order. The tree can't see steps wound up by hand.", await page.InnerTextAsync(".calls-note"));

    // From Widget.Stop: the step it calls, which a call can't end at, as another extends it; and a call it can't follow.
    await page.Locator("#calls-tree li[data-id='Widget.Spin'] .calls-name").ClickAsync(new() { Modifiers = [KeyboardModifier.Control] });
    await opened.SelectAsync(null);
    await page.EvaluateAsync("() => { window.codeViews.select('Widget.Stop'); }");
    await page.Locator("#calls-tree li[data-id='Widget.Stop']").First.WaitForAsync();
    var text = await page.EvaluateAsync<string>("() => window.codeViewsCalls.text()");
    if (!text.Contains("Blueprint.Draw", StringComparison.Ordinal))
    {
      await page.Locator("#calls-tree li[data-id='Widget.Stop'] > .tree-row .tree-twisty").First.ClickAsync();
    }

    await page.Locator("#calls-tree li[data-id='Blueprint.Draw'] > .tree-row .tree-twisty").First.ClickAsync();
    text = await page.EvaluateAsync<string>("() => window.codeViewsCalls.text()");
    Assert.Contains(string.Join("\n",
      "- Widget.Stop (Step)",
      "  - Blueprint.Draw (Step), line 5",
      "    - overridden by: Gizmo.Run (Step)",
      "  - `thing.poke` (loose), line 6"), text);
    Assert.Equal("Wound up by hand, so what it runs isn't written down", await page.GetAttributeAsync("#calls-tree li.calls-late .tree-detail", "title"));
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task The_side_panel_names_kinds_roles_and_links_as_the_language_does()
  {
    await using var opened = await fixture.OpenAsync(fixture.Embedded);
    var page = opened.Page;

    await opened.SelectAsync("Gizmo");
    Assert.Equal("Thing", await page.InnerTextAsync("#side-panel-header h2 .kind"));
    await page.ClickAsync("#source-dependencies > summary");
    Assert.Equal(["extends"], await Texts(page, "#source-dependencies [data-direction='uses'] .edge-tag"));
    Assert.Equal("Extends this", await page.GetAttributeAsync("#source-dependencies [data-direction='uses'] .edge-tag", "title"));

    await opened.SelectAsync("Blueprint");
    Assert.Equal(["extended by"], await Texts(page, "#source-dependencies [data-direction='usedBy'] .edge-tag"));

    await opened.SelectAsync("Widget.Spin");
    Assert.Equal("Step in Widget", await page.InnerTextAsync("#side-panel-header h2 .kind"));
    Assert.Equal(["Starting point"], await Texts(page, "#side-panel-header .badge"));

    await opened.SelectAsync("Widget.Stop");
    Assert.Contains("Loose, so not followed:", await page.InnerTextAsync("#side-panel-header"));
    Assert.Equal("Show in the whole gadget", await page.InnerTextAsync("#toggle-context"));
    Assert.Empty(opened.Errors);
  }
}
