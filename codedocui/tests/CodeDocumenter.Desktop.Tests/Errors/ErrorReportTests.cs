using System.Runtime.CompilerServices;
using CodeDocumenter.Desktop.Errors;
using CodeDocumenter.Shared;

namespace CodeDocumenter.Desktop.Tests.Errors;

public sealed class ErrorReportTests
{
  private const string SecretPath = @"C:\Users\someone\Payroll\Salaries 2026.xlsm";

  private static readonly ErrorContext Context = new(
    new DateTimeOffset(2026, 11, 17, 14, 30, 12, TimeSpan.FromHours(11)),
    [new ToolVersion("codedocui", "0.8.0", "abc1234def", Dirty: false), new ToolVersion("vbadoc", "0.43.0", null, Dirty: false)],
    OnWindowThread: true,
    ["Tab showing: Import", "Import tab: source a folder"]);

  // Thrown from a method of this file, so the stack trace has this file's path in it, as the app's have theirs.
  private static Exception Thrown(Func<Exception> make)
  {
    try
    {
      throw make();
    }
    catch (Exception ex)
    {
      return ex;
    }
  }

  private static string ThisFolder([CallerFilePath] string path = "") => Path.GetDirectoryName(path)!;

  [Fact]
  public void A_report_has_no_message_and_no_folder_but_has_the_error_and_where_it_came_from()
  {
    var exception = Thrown(() => new IOException($"The process cannot access the file '{SecretPath}'."));

    var report = ErrorReport.Write(exception, Context);

    Assert.DoesNotContain("Payroll", report, StringComparison.OrdinalIgnoreCase);
    Assert.DoesNotContain("someone", report, StringComparison.OrdinalIgnoreCase);
    Assert.DoesNotContain(ThisFolder(), report, StringComparison.OrdinalIgnoreCase);
    Assert.Contains("System.IO.IOException", report, StringComparison.Ordinal);
    Assert.Contains($"{nameof(ErrorReportTests)}.{nameof(Thrown)}", report, StringComparison.Ordinal);
    Assert.Contains($" in {nameof(ErrorReportTests)}.cs:", report, StringComparison.Ordinal);
  }

  [Fact]
  public void A_report_says_which_build_on_what_and_what_the_app_was_set_to()
  {
    var report = ErrorReport.Write(Thrown(() => new InvalidOperationException()), Context);

    Assert.StartsWith("codedocui error report\r\n", report, StringComparison.Ordinal);
    Assert.Contains("When:    2026-11-17 14:30:12 +11:00", report, StringComparison.Ordinal);
    Assert.Contains("Build:   codedocui 0.8.0 (abc1234) · vbadoc 0.43.0", report, StringComparison.Ordinal);
    Assert.Contains(".NET", report, StringComparison.Ordinal);
    Assert.Contains("Thread:  the window's", report, StringComparison.Ordinal);
    Assert.Contains("  Tab showing: Import\r\n  Import tab: source a folder\r\n", report, StringComparison.Ordinal);
  }

  [Fact]
  public void Errors_inside_an_error_are_reported_with_theirs_and_without_their_messages()
  {
    var inner = new FileNotFoundException($"Could not find file '{SecretPath}'.", SecretPath);
    var exception = Thrown(() => new InvalidOperationException("Reading failed.", inner));
    var aggregate = new AggregateException(exception, new UnauthorizedAccessException(SecretPath));

    var report = ErrorReport.Write(aggregate, Context);

    Assert.Contains("System.AggregateException", report, StringComparison.Ordinal);
    Assert.Contains("  System.InvalidOperationException", report, StringComparison.Ordinal);
    Assert.Contains("    System.IO.FileNotFoundException", report, StringComparison.Ordinal);
    Assert.Contains("  System.UnauthorizedAccessException", report, StringComparison.Ordinal);
    Assert.DoesNotContain("Payroll", report, StringComparison.OrdinalIgnoreCase);
    Assert.DoesNotContain("Reading failed", report, StringComparison.Ordinal);
  }

  [Theory]
  [InlineData(@"   at A.B() in C:\Users\someone\repo\src\File.cs:line 12", "   at A.B() in File.cs:line 12")]
  [InlineData(@"   at A.B() in \\server\share\src\File.cs:line 3", "   at A.B() in File.cs:line 3")]
  [InlineData("   at A.B() in /_/src/File.cs:line 4", "   at A.B() in File.cs:line 4")]
  [InlineData("   at A.B() in /home/someone/src/File.cs:line 5", "   at A.B() in File.cs:line 5")]
  [InlineData(@"   at A.B() in C:\Program Files (x86)\Some Thing\File.cs:Zeile 6", "   at A.B() in File.cs:Zeile 6")]
  [InlineData("   at System.Linq.Enumerable.Select[TSource,TResult](IEnumerable`1 source, Func`2 selector)", "   at System.Linq.Enumerable.Select[TSource,TResult](IEnumerable`1 source, Func`2 selector)")]
  [InlineData("--- End of stack trace from previous location ---", "--- End of stack trace from previous location ---")]
  [InlineData("   at CodeDocumenter.Desktop.Window.ImportTab.<>c.<ShowSummary>b__58_0(ListViewItem item)", "   at CodeDocumenter.Desktop.Window.ImportTab.<>c.<ShowSummary>b__58_0(ListViewItem item)")]
  public void A_stack_traces_paths_are_cut_to_their_files_names(string line, string cleaned)
  {
    Assert.Equal(new[] { cleaned }, ErrorReport.Clean(line + "\r\n"));
  }

  [Fact]
  public void An_error_never_thrown_has_no_stack_trace_and_says_so()
  {
    Assert.Equal(new[] { "   (no stack trace)" }, ErrorReport.Clean(new InvalidOperationException().StackTrace));
  }

  [Fact]
  public void The_summary_says_what_the_error_was_and_the_file_and_line_it_came_from()
  {
    var summary = ErrorReport.Summary(Thrown(() => new IOException(SecretPath)));

    Assert.Matches($@"^IOException in {nameof(ErrorReportTests)}\.cs, line \d+$", summary);
    Assert.Equal("ArgumentException", ErrorReport.Summary(new ArgumentException(SecretPath)));
  }
}
