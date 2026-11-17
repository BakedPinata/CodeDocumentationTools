using CodeDocumenter.Desktop.Errors;
using CodeDocumenter.Shared;

namespace CodeDocumenter.Desktop.Tests.Errors;

public sealed class ErrorHandlerTests : IDisposable
{
  private readonly TempFolder folder = new();
  private readonly FakeDialog dialog = new();
  private readonly List<int> exits = [];
  private int startedAfresh;

  public void Dispose() => folder.Dispose();

  private string Reports => folder.Combine("reports");

  private ErrorHandler Handler(Func<IReadOnlyList<string>>? describeState = null, string? reports = null) => new(
    reports ?? Reports,
    [new ToolVersion("codedocui", "0.8.0", null, Dirty: false)],
    describeState ?? (() => ["Tab showing: Import"]),
    dialog,
    () => startedAfresh++,
    exits.Add);

  [Fact]
  public void An_error_is_saved_and_shown_once_and_then_the_app_closes()
  {
    Handler().Handle(new InvalidOperationException("C:\\Secret\\Book.xlsm"), onWindowThread: true);

    var shown = Assert.Single(dialog.Shown);
    Assert.Equal("InvalidOperationException", shown.Summary);
    Assert.Equal(Assert.Single(Directory.GetFiles(Reports)), shown.ReportPath);
    Assert.Equal(shown.Report, File.ReadAllText(shown.ReportPath!));
    Assert.Contains("Tab showing: Import", shown.Report, StringComparison.Ordinal);
    Assert.DoesNotContain("Secret", shown.Report, StringComparison.Ordinal);
    Assert.Equal(new[] { ExitCodes.UnexpectedError }, exits);
    Assert.Equal(0, startedAfresh);
  }

  [Fact]
  public void Errors_while_one_is_showing_are_counted_not_shown()
  {
    var handler = Handler();
    dialog.WhileShowing = () =>
    {
      handler.Handle(new NullReferenceException(), onWindowThread: true);
      handler.Handle(new NullReferenceException(), onWindowThread: false);
    };

    handler.Handle(new NullReferenceException(), onWindowThread: true);

    var shown = Assert.Single(dialog.Shown);
    Assert.EndsWith("Errors after it, while it was showing: 2\r\n", File.ReadAllText(shown.ReportPath!), StringComparison.Ordinal);
    Assert.Equal(new[] { ExitCodes.UnexpectedError }, exits);
  }

  [Fact]
  public void Starting_afresh_puts_the_settings_aside_before_the_app_closes()
  {
    dialog.StartAfresh = true;

    Handler().Handle(new InvalidOperationException(), onWindowThread: true);

    Assert.Equal(1, startedAfresh);
    Assert.Equal(new[] { ExitCodes.UnexpectedError }, exits);
  }

  [Fact]
  public void A_report_that_cant_be_saved_is_still_shown()
  {
    var notAFolder = folder.Combine("a file");
    File.WriteAllText(notAFolder, "");

    Handler(reports: notAFolder).Handle(new InvalidOperationException(), onWindowThread: true);

    var shown = Assert.Single(dialog.Shown);
    Assert.Null(shown.ReportPath);
    Assert.Contains("System.InvalidOperationException", shown.Report, StringComparison.Ordinal);
    Assert.Equal(new[] { ExitCodes.UnexpectedError }, exits);
  }

  [Fact]
  public void A_window_that_cant_say_what_it_was_set_to_still_gets_a_report()
  {
    Handler(describeState: () => throw new ObjectDisposedException("window")).Handle(new InvalidOperationException(), onWindowThread: true);

    Assert.Contains("(couldn't be read: ObjectDisposedException)", Assert.Single(dialog.Shown).Report, StringComparison.Ordinal);
  }

  [Fact]
  public void A_dialog_that_fails_still_closes_the_app()
  {
    dialog.WhileShowing = () => throw new InvalidOperationException("No dialog today.");

    Handler().Handle(new NullReferenceException(), onWindowThread: true);

    Assert.Equal(new[] { ExitCodes.UnexpectedError }, exits);
  }

  private sealed class FakeDialog : IErrorDialog
  {
    public List<ShownError> Shown { get; } = [];

    public bool StartAfresh { get; set; }

    public Action? WhileShowing { get; set; }

    public bool Show(ShownError error)
    {
      Shown.Add(error);
      WhileShowing?.Invoke();
      return StartAfresh;
    }
  }
}
