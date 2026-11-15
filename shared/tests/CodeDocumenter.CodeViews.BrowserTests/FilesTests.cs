using Microsoft.Playwright;

namespace CodeDocumenter.CodeViews.BrowserTests;

/// <summary>
/// Files are apart from what's in them (step M3 of the generalisation plan): a unit can be in two files, and a file
/// can hold two units. Each unit and member says which file and lines it's in, and the source panel shows those.
/// </summary>
[Collection("Browser")]
public sealed class FilesTests(ToyPageFixture fixture)
{
  [Fact]
  public async Task A_unit_in_two_files_shows_one_part_at_a_time_and_says_where_each_is()
  {
    await using var opened = await fixture.OpenAsync(fixture.Embedded);
    var page = opened.Page;
    await opened.SelectAsync("Widget");

    Assert.Equal("a.toy · lines 1 to 6; b.toy · lines 3 to 8", await page.InnerTextAsync("#side-panel-header .signature"));
    Assert.Equal(["a.toy", "b.toy"], await page.Locator("#side-panel-header [data-part]").AllInnerTextsAsync());
    Assert.Equal(new[] { 1, 2, 3, 4, 5, 6 }, await opened.ShownLinesAsync());
    Assert.Equal("gadget Widget", await page.InnerTextAsync("#side-panel-body .line >> nth=0 >> .lt"));

    await page.ClickAsync("#side-panel-header [data-part='1']");
    Assert.Equal(new[] { 3, 4, 5, 6, 7, 8 }, await opened.ShownLinesAsync());
    Assert.Equal("more Widget", await page.InnerTextAsync("#side-panel-body .line >> nth=0 >> .lt"));
    Assert.Equal("true", await page.GetAttributeAsync("#side-panel-header [data-part='1']", "aria-pressed"));

    // Another selection starts from the first part again.
    await opened.SelectAsync("Blueprint");
    await opened.SelectAsync("Widget");
    Assert.Equal(new[] { 1, 2, 3, 4, 5, 6 }, await opened.ShownLinesAsync());
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task A_member_shows_its_own_file_and_lines_whichever_part_of_its_unit_holds_it()
  {
    await using var opened = await fixture.OpenAsync(fixture.Embedded);
    var page = opened.Page;

    await opened.SelectAsync("Widget.Stop");
    Assert.Equal(new[] { 4, 5, 6, 7 }, await opened.ShownLinesAsync());
    Assert.Equal("step Stop", (await page.InnerTextAsync("#side-panel-body .line >> nth=0 >> .lt")).Trim());

    // In its whole unit: the part that holds it, with its lines marked, not the rest of the file.
    await page.ClickAsync("#toggle-context");
    Assert.Equal(new[] { 3, 4, 5, 6, 7, 8 }, await opened.ShownLinesAsync());
    Assert.Equal(["4", "5", "6", "7"], await page.Locator("#side-panel-body .line.marked .ln").AllInnerTextsAsync());
    Assert.Equal("Show this part only", await page.InnerTextAsync("#toggle-context"));

    await opened.SelectAsync("Widget.Spin");
    Assert.Equal(new[] { 2, 3, 4 }, await opened.ShownLinesAsync());
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task A_file_with_two_units_shows_each_units_own_lines()
  {
    await using var opened = await fixture.OpenAsync(fixture.Embedded);
    var page = opened.Page;

    await opened.SelectAsync("Gizmo");
    Assert.Equal("b.toy · lines 10 to 14", await page.InnerTextAsync("#side-panel-header .signature"));
    Assert.Equal(0, await page.Locator("#side-panel-header [data-part]").CountAsync());
    Assert.Equal(new[] { 10, 11, 12, 13, 14 }, await opened.ShownLinesAsync());

    // A unit that's the whole of its file says how long the file is.
    await opened.SelectAsync("Blueprint");
    Assert.Equal("c.toy · 5 lines", await page.InnerTextAsync("#side-panel-header .signature"));
    Assert.Equal(new[] { 1, 2, 3, 4, 5 }, await opened.ShownLinesAsync());
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task A_reference_is_shown_in_the_file_of_the_member_that_makes_it()
  {
    await using var opened = await fixture.OpenAsync(fixture.Embedded);
    var page = opened.Page;
    await opened.SelectAsync("Blueprint.Draw");
    await page.ClickAsync("#source-dependencies > summary");

    // Used by Widget.Stop, on line 5 of b.toy, which isn't the first file of Widget.
    await page.Locator("#source-dependencies [data-direction='usedBy'] details > summary").First.ClickAsync();
    var line = page.Locator("#source-dependencies [data-direction='usedBy'] [data-peek]").First;
    Assert.Contains("Blueprint.Draw", await line.InnerTextAsync());
    await line.ClickAsync(new() { Modifiers = [KeyboardModifier.Control] });

    Assert.Equal("Widget.Stop", await page.EvaluateAsync<string>("() => window.codeViews.selection()"));
    Assert.Equal(["5"], await page.Locator("#side-panel-body .line.target").EvaluateAllAsync<string[]>("ls => ls.map(l => l.dataset.line)"));
    Assert.Equal("    Blueprint.Draw", await page.Locator("#side-panel-body .line.target .lt").TextContentAsync());
    Assert.Empty(opened.Errors);
  }

  [Fact]
  public async Task Load_folder_reads_the_files_by_name_and_shows_a_unit_from_each()
  {
    await using var opened = await fixture.OpenAsync(fixture.Folder);
    var page = opened.Page;
    await opened.SelectAsync("Widget.Stop");
    Assert.Contains("contains this page and its .toy files", await page.InnerTextAsync("#side-panel-body"));

    await page.SetInputFilesAsync("#load-folder", fixture.Folder);

    await page.WaitForFunctionAsync("() => document.getElementById('source-status-text').textContent.startsWith('Loaded')");
    Assert.Equal("Loaded 3 file(s)", await page.InnerTextAsync("#source-status-text"));
    Assert.Equal(new[] { 4, 5, 6, 7 }, await opened.ShownLinesAsync());
    await opened.SelectAsync("Blueprint");
    Assert.Equal(new[] { 1, 2, 3, 4, 5 }, await opened.ShownLinesAsync());
    Assert.Empty(opened.Errors);
  }
}
