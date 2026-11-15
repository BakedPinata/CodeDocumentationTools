using Microsoft.Playwright;

namespace CodeDocumenter.CodeViews.BrowserTests;

/// <summary>
/// Shared by the browser tests: installs Playwright's Chromium on first use (a one-off download), starts it headless,
/// and writes the made-up language's page to a temporary folder.
/// </summary>
public sealed class ToyPageFixture : IAsyncLifetime
{
  private readonly string root = Path.Combine(Path.GetTempPath(), $"codeviews-browser-{Guid.NewGuid():N}");

  public IPlaywright Playwright { get; private set; } = null!;

  public IBrowser Browser { get; private set; } = null!;

  /// <summary>The page with its source embedded.</summary>
  public string Embedded { get; private set; } = string.Empty;

  /// <summary>The page in a folder with its source files, for Load folder.</summary>
  public string Folder { get; private set; } = string.Empty;

  public async ValueTask InitializeAsync()
  {
    var exitCode = Microsoft.Playwright.Program.Main(["install", "chromium"]);
    if (exitCode != 0)
    {
      throw new InvalidOperationException($"Installing Chromium for Playwright failed (exit code {exitCode}).");
    }

    Playwright = await Microsoft.Playwright.Playwright.CreateAsync();
    Browser = await Playwright.Chromium.LaunchAsync();

    Embedded = Path.Combine(root, "embedded");
    Directory.CreateDirectory(Embedded);
    File.WriteAllText(Path.Combine(Embedded, "Toybox-toy.html"), Toy.Page(embedSource: true));

    Folder = Path.Combine(root, "folder");
    Directory.CreateDirectory(Folder);
    File.WriteAllText(Path.Combine(Folder, "Toybox-toy.html"), Toy.Page(embedSource: false));
    foreach (var (name, text) in Toy.Sources)
    {
      File.WriteAllText(Path.Combine(Folder, name), text);
    }
  }

  public async ValueTask DisposeAsync()
  {
    await Browser.DisposeAsync();
    Playwright.Dispose();
    if (Directory.Exists(root))
    {
      Directory.Delete(root, recursive: true);
    }
  }

  /// <summary>Opens a page, keeping what it logs as errors.</summary>
  public async Task<OpenedToyPage> OpenAsync(string folder, ColorScheme colourScheme = ColorScheme.Dark)
  {
    var context = await Browser.NewContextAsync(new() { ColorScheme = colourScheme, ViewportSize = new() { Width = 1280, Height = 800 } });
    var page = await context.NewPageAsync();
    var errors = new List<string>();
    page.Console += (_, message) => { if (message.Type == "error") errors.Add(message.Text); };
    page.PageError += (_, error) => errors.Add(error);
    await page.GotoAsync(new Uri(Path.Combine(folder, "Toybox-toy.html")).AbsoluteUri);
    await page.WaitForFunctionAsync("() => !!window.codeViews && !!window.codeViewsGraph");
    return new OpenedToyPage(context, page, errors);
  }
}

[CollectionDefinition("Browser")]
public sealed class BrowserCollection : ICollectionFixture<ToyPageFixture>;

/// <summary>An open page, and the errors it logged.</summary>
public sealed class OpenedToyPage(IBrowserContext context, IPage page, List<string> errors) : IAsyncDisposable
{
  public IPage Page { get; } = page;

  /// <summary>Console errors and uncaught exceptions.</summary>
  public List<string> Errors { get; } = errors;

  public Task SelectAsync(string? id) => Page.EvaluateAsync("id => window.codeViews.select(id)", id);

  /// <summary>The line numbers the source panel shows.</summary>
  public async Task<int[]> ShownLinesAsync() =>
    [.. (await Page.Locator("#side-panel-body .line .ln").AllInnerTextsAsync()).Select(int.Parse)];

  public async ValueTask DisposeAsync() => await context.DisposeAsync();
}
