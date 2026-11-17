using System.Diagnostics;

namespace CodeDocumenter.Shared.Files;

/// <summary>
/// Waiting out another program that has a file open for a moment, as endpoint protection, a sync client or Windows Search
/// often do just after a file is read or written. Only Windows saying the file is in use is waited for, and only for so
/// long: any other error, or the same one once the wait is up, is thrown as it came. Nothing has been done to a file
/// Windows says is in use, so trying again is safe.
/// </summary>
public static class FileInUse
{
  // Windows' ERROR_SHARING_VIOLATION and ERROR_LOCK_VIOLATION, as an IOException's HResult.
  private const int SharingViolation = unchecked((int)0x80070020);
  private const int LockViolation = unchecked((int)0x80070021);

  private static readonly TimeSpan FirstPause = TimeSpan.FromMilliseconds(50);
  private static readonly TimeSpan LongestPause = TimeSpan.FromMilliseconds(500);

  /// <summary>How long a file in use is waited for, unless the caller says otherwise.</summary>
  public static TimeSpan Wait { get; } = TimeSpan.FromSeconds(5);

  /// <summary>Whether an error is Windows saying another program has the file open.</summary>
  public static bool Is(Exception exception) => exception is IOException { HResult: SharingViolation or LockViolation };

  /// <summary>
  /// Does <paramref name="act"/>, and while the file's in use, does it again after a pause, each twice as long as the one
  /// before up to half a second, until <paramref name="wait"/> is up.
  /// </summary>
  /// <param name="act">What to do with the file.</param>
  /// <param name="wait">How long to keep trying; <see cref="Wait"/> if null.</param>
  /// <param name="pausing">Told of each pause and how long it is, for a log.</param>
  public static T Retry<T>(Func<T> act, TimeSpan? wait = null, Action<TimeSpan>? pausing = null)
  {
    ArgumentNullException.ThrowIfNull(act);
    var until = wait ?? Wait;
    var waited = Stopwatch.StartNew();
    var pause = FirstPause;
    while (true)
    {
      try
      {
        return act();
      }
      catch (IOException ex) when (Is(ex) && waited.Elapsed + pause <= until)
      {
        pausing?.Invoke(pause);
        Thread.Sleep(pause);
        pause = pause * 2 < LongestPause ? pause * 2 : LongestPause;
      }
    }
  }

  /// <inheritdoc cref="Retry{T}(Func{T}, TimeSpan?, Action{TimeSpan}?)" />
  public static void Retry(Action act, TimeSpan? wait = null, Action<TimeSpan>? pausing = null)
  {
    ArgumentNullException.ThrowIfNull(act);
    Retry(() =>
    {
      act();
      return true;
    }, wait, pausing);
  }
}
