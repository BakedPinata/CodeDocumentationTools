using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using CodeDocumenter.Shared;

namespace CodeDocumenter.Desktop.Errors;

/// <summary>What's known when an error comes: when, which build, on which thread, and what the app was set to.</summary>
/// <param name="When">When it came.</param>
/// <param name="Versions">The app's version and the tools'.</param>
/// <param name="OnWindowThread">Whether it came on the window's thread, rather than another.</param>
/// <param name="State">What the app was set to, in kinds and counts only (<see cref="StateSummary"/>).</param>
internal sealed record ErrorContext(DateTimeOffset When, IReadOnlyList<ToolVersion> Versions, bool OnWindowThread, IReadOnlyList<string> State);

/// <summary>
/// A report of an error the app didn't expect, to read and to pass on. It says which build it was, on what, what the app
/// was set to and where in the code the error came from, and nothing of yours: no error's message, since a file's error
/// names its path; no file's folder, including the folder the app was built in, which can name whoever built it; and of
/// what the app was set to, only kinds and counts.
/// </summary>
internal static partial class ErrorReport
{
  // Errors inside errors are followed this deep, which is deeper than any the app makes.
  private const int MostDepth = 8;

  /// <summary>The report's text.</summary>
  public static string Write(Exception exception, ErrorContext context)
  {
    ArgumentNullException.ThrowIfNull(exception);
    ArgumentNullException.ThrowIfNull(context);
    var text = new StringBuilder();
    Line(text, "codedocui error report");
    Line(text, "");
    Line(text, "It names no file or folder and says nothing of what's in them, so it can be passed on as it is.");
    Line(text, "");
    Line(text, $"When:    {context.When.ToString("yyyy-MM-dd HH:mm:ss zzz", CultureInfo.InvariantCulture)}");
    Line(text, $"Build:   {string.Join(" · ", context.Versions)}");
    Line(text, $"On:      {RuntimeInformation.FrameworkDescription}, {RuntimeInformation.OSDescription}, {RuntimeInformation.ProcessArchitecture}");
    Line(text, $"Thread:  {(context.OnWindowThread ? "the window's" : "another")}");
    Line(text, "");
    Line(text, "What the app was set to:");
    foreach (var state in context.State)
    {
      Line(text, $"  {state}");
    }

    Line(text, "");
    Line(text, "The error:");
    AppendError(text, exception, "");
    return text.ToString();
  }

  /// <summary>
  /// A line saying what the error was and where it came from, such as "NullReferenceException in ImportTab.cs, line 446",
  /// with no message, for the dialog.
  /// </summary>
  public static string Summary(Exception exception)
  {
    ArgumentNullException.ThrowIfNull(exception);
    var frame = new StackTrace(exception, fNeedFileInfo: true).GetFrames().FirstOrDefault(f => f.GetFileName() is not null);
    return frame?.GetFileName() is { } file
      ? string.Create(CultureInfo.InvariantCulture, $"{exception.GetType().Name} in {Path.GetFileName(file)}, line {frame.GetFileLineNumber()}")
      : exception.GetType().Name;
  }

  /// <summary>A stack trace's lines, with each path cut to the file's name, or a line saying there's none.</summary>
  public static IReadOnlyList<string> Clean(string? stackTrace) => string.IsNullOrWhiteSpace(stackTrace)
    ? ["   (no stack trace)"]
    : [.. stackTrace.Split('\n').Select(line => Folders().Replace(line.TrimEnd('\r'), "")).Where(line => line.Length > 0)];

  private static void AppendError(StringBuilder text, Exception exception, string indent)
  {
    Line(text, $"{indent}{exception.GetType().FullName} (HResult 0x{exception.HResult:X8})");
    foreach (var line in Clean(exception.StackTrace))
    {
      Line(text, indent + line);
    }

    if (indent.Length / 2 >= MostDepth)
    {
      return;
    }

    IEnumerable<Exception> inside = exception is AggregateException aggregate
      ? aggregate.InnerExceptions
      : exception.InnerException is { } inner ? [inner] : [];
    foreach (var each in inside)
    {
      Line(text, $"{indent}Inside it:");
      AppendError(text, each, indent + "  ");
    }
  }

  // Lines end the same on every machine, so a report reads the same wherever it's opened.
  private static void Line(StringBuilder text, string line) => text.Append(line).Append("\r\n");

  // The folders of a path, Windows (C:\…\ or \\server\share\…\) or not (/…/), up to its last separator, so only the file's
  // name is left. No part of a stack trace but a path has a slash in it.
  [GeneratedRegex(@"(?:[A-Za-z]:)?(?:[\\/][^\\/:*?""<>|\r\n]*)*[\\/]", RegexOptions.CultureInvariant)]
  private static partial Regex Folders();
}
