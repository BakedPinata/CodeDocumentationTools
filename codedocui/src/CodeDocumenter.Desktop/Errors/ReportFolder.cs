using System.Globalization;
using System.Text;
using CodeDocumenter.Desktop.Settings;

namespace CodeDocumenter.Desktop.Errors;

/// <summary>Where error reports are kept: text files beside the settings, the newest twenty.</summary>
internal static class ReportFolder
{
  /// <summary>How many reports are kept. They're small, but an app that keeps failing shouldn't fill a disk.</summary>
  public const int Kept = 20;

  private const string Prefix = "codedocui-error-";

  /// <summary><c>%LOCALAPPDATA%\codedocui\reports</c>.</summary>
  public static string DefaultPath { get; } = Path.Combine(Path.GetDirectoryName(SettingsStore.DefaultPath)!, "reports");

  /// <summary>
  /// Saves a report as <c>codedocui-error-&lt;when&gt;.txt</c> and lets the oldest go beyond <see cref="Kept"/>.
  /// </summary>
  /// <returns>Where it was saved.</returns>
  /// <exception cref="IOException">It couldn't be saved.</exception>
  /// <exception cref="UnauthorizedAccessException">It couldn't be saved.</exception>
  public static string Save(string folder, string report, DateTimeOffset when)
  {
    ArgumentNullException.ThrowIfNull(folder);
    ArgumentNullException.ThrowIfNull(report);
    Directory.CreateDirectory(folder);
    var name = Prefix + when.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
    var path = Path.Combine(folder, name + ".txt");
    for (var n = 2; File.Exists(path); n++)
    {
      path = Path.Combine(folder, string.Create(CultureInfo.InvariantCulture, $"{name}-{n}.txt"));
    }

    File.WriteAllText(path, report, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    LetOldOnesGo(folder);
    return path;
  }

  // The newest are those written last, whatever the clock said when each was named.
  private static void LetOldOnesGo(string folder)
  {
    var old = Directory.GetFiles(folder, Prefix + "*.txt")
      .OrderByDescending(path => File.GetLastWriteTimeUtc(path))
      .ThenByDescending(path => path, StringComparer.OrdinalIgnoreCase)
      .Skip(Kept);
    foreach (var path in old)
    {
      try
      {
        File.Delete(path);
      }
      catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
      {
        // Left for next time: the report just saved matters more.
      }
    }
  }
}
