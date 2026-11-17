using System.Globalization;
using CodeDocumenter.Shared;

namespace CodeDocumenter.Desktop.Errors;

/// <summary>An error as it's shown: a line on what it was, its report, and where that was saved, or null if it couldn't be.</summary>
internal sealed record ShownError(string Summary, string Report, string? ReportPath);

/// <summary>Shows an error and its report, and says whether to start afresh next time.</summary>
internal interface IErrorDialog
{
  /// <returns>Whether the settings are to be put aside, so the app starts afresh next time.</returns>
  bool Show(ShownError error);
}

/// <summary>
/// What happens on an error the app didn't expect. Carrying on could bring the same error again and again, so the first
/// is reported and the app closes: its report is saved and shown, and when that's closed, so is the app, without
/// saving the settings over the ones it had. Errors that come while it's showing are counted in the report rather than
/// shown.
/// </summary>
/// <param name="reportFolder">Where reports are saved.</param>
/// <param name="versions">The app's version and the tools'.</param>
/// <param name="describeState">What the app was set to, in kinds and counts only (<see cref="StateSummary"/>).</param>
/// <param name="dialog">Shows the error.</param>
/// <param name="startAfresh">Puts the settings aside.</param>
/// <param name="exit">Ends the app with an exit code. In the app it doesn't return.</param>
/// <param name="time">The clock, for the report's time.</param>
internal sealed class ErrorHandler(
  string reportFolder,
  IReadOnlyList<ToolVersion> versions,
  Func<IReadOnlyList<string>> describeState,
  IErrorDialog dialog,
  Action startAfresh,
  Action<int> exit,
  TimeProvider? time = null)
{
  private readonly TimeProvider time = time ?? TimeProvider.System;
  private int handling;
  private int repeats;

  /// <summary>Reports the error and ends the app, or, if another is being reported, counts it.</summary>
  public void Handle(Exception exception, bool onWindowThread)
  {
    ArgumentNullException.ThrowIfNull(exception);
    if (Interlocked.Exchange(ref handling, 1) == 1)
    {
      Interlocked.Increment(ref repeats);
      return;
    }

    var when = time.GetLocalNow();
    var report = ErrorReport.Write(exception, new ErrorContext(when, versions, onWindowThread, State()));
    var path = Save(report, when);
    var afresh = Show(new ShownError(ErrorReport.Summary(exception), report, path));
    var more = Volatile.Read(ref repeats);
    if (path is not null && more > 0)
    {
      Try(() => File.AppendAllText(path, string.Create(CultureInfo.InvariantCulture, $"\r\nErrors after it, while it was showing: {more}\r\n")));
    }

    if (afresh)
    {
      Try(startAfresh);
    }

    exit(ExitCodes.UnexpectedError);
  }

  // What the app was set to, if it can be read: the error may have left the window unable to say.
  private IReadOnlyList<string> State()
  {
    try
    {
      return describeState();
    }
    catch (Exception ex)
    {
      return [$"(couldn't be read: {ex.GetType().Name})"];
    }
  }

  private string? Save(string report, DateTimeOffset when)
  {
    try
    {
      return ReportFolder.Save(reportFolder, report, when);
    }
    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
    {
      // The dialog has the report to copy instead.
      return null;
    }
  }

  // Whatever the dialog does, the app still closes.
  private bool Show(ShownError error)
  {
    try
    {
      return dialog.Show(error);
    }
    catch (Exception)
    {
      return false;
    }
  }

  private static void Try(Action action)
  {
    try
    {
      action();
    }
    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
    {
      // Closing matters more.
    }
  }
}
