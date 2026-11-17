using System.Diagnostics;
using CodeDocumenter.Shared.Files;

namespace CodeDocumenter.Shared.Tests.Files;

public sealed class FileInUseTests : IDisposable
{
  private static readonly IOException InUse = new("The process cannot access the file because it is being used by another process.", unchecked((int)0x80070020));

  private readonly TempDirectory temp = new();

  public void Dispose() => temp.Dispose();

  [Fact]
  public void Only_a_file_in_use_counts()
  {
    Assert.True(FileInUse.Is(InUse));
    Assert.True(FileInUse.Is(new IOException("locked", unchecked((int)0x80070021))));
    Assert.False(FileInUse.Is(new FileNotFoundException()));
    Assert.False(FileInUse.Is(new UnauthorizedAccessException()));
  }

  [Fact]
  public void A_file_in_use_is_tried_again_after_a_pause_each_longer_than_the_one_before()
  {
    var tries = 0;
    var pauses = new List<TimeSpan>();

    var result = FileInUse.Retry(() => ++tries < 4 ? throw InUse : "done", pausing: pauses.Add);

    Assert.Equal("done", result);
    Assert.Equal(4, tries);
    Assert.Equal(new[] { 50, 100, 200 }, pauses.Select(p => (int)p.TotalMilliseconds));
  }

  [Fact]
  public void Any_other_error_is_thrown_at_once()
  {
    var tries = 0;

    Assert.Throws<FileNotFoundException>(() => FileInUse.Retry(() =>
    {
      tries++;
      throw new FileNotFoundException();
    }));

    Assert.Equal(1, tries);
  }

  [Fact]
  public void A_file_still_in_use_when_the_wait_is_up_is_thrown_as_it_came()
  {
    var timer = Stopwatch.StartNew();

    var thrown = Assert.Throws<IOException>(() => FileInUse.Retry(() => throw InUse, TimeSpan.FromMilliseconds(300)));

    Assert.Same(InUse, thrown);
    Assert.InRange(timer.Elapsed, TimeSpan.FromMilliseconds(100), TimeSpan.FromSeconds(3));
  }

  [Fact]
  public void A_file_another_program_lets_go_of_is_had_once_it_has()
  {
    Assert.SkipUnless(OperatingSystem.IsWindows(), "Only Windows stops a file being opened that another program has open.");
    var path = temp.WriteFile("held.txt", "held");
    using var holder = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
    using var letGo = new Timer(_ => holder.Dispose(), null, 300, Timeout.Infinite);

    var text = FileInUse.Retry(() => File.ReadAllText(path));

    Assert.Equal("held", text);
  }
}
