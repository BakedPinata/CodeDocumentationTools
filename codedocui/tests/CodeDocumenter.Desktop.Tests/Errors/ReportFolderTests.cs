using CodeDocumenter.Desktop.Errors;

namespace CodeDocumenter.Desktop.Tests.Errors;

public sealed class ReportFolderTests : IDisposable
{
  private static readonly DateTimeOffset When = new(2026, 11, 17, 14, 30, 12, TimeSpan.FromHours(11));

  private readonly TempFolder folder = new();

  private string Reports => folder.Combine("codedocui", "reports");

  public void Dispose() => folder.Dispose();

  [Fact]
  public void A_report_is_saved_named_for_when_it_came()
  {
    var path = ReportFolder.Save(Reports, "the report", When);

    Assert.Equal(Path.Combine(Reports, "codedocui-error-20261117-143012.txt"), path);
    Assert.Equal("the report", File.ReadAllText(path));
  }

  [Fact]
  public void Two_in_the_same_second_are_both_kept()
  {
    var first = ReportFolder.Save(Reports, "first", When);
    var second = ReportFolder.Save(Reports, "second", When);

    Assert.Equal("first", File.ReadAllText(first));
    Assert.Equal(Path.Combine(Reports, "codedocui-error-20261117-143012-2.txt"), second);
    Assert.Equal("second", File.ReadAllText(second));
  }

  [Fact]
  public void Only_the_newest_are_kept()
  {
    // Written a minute apart, long ago, so the one each save makes now is the newest.
    var longAgo = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
    var paths = Enumerable.Range(0, ReportFolder.Kept + 2)
      .Select(minutes =>
      {
        var path = ReportFolder.Save(Reports, $"report {minutes}", When.AddMinutes(minutes));
        File.SetLastWriteTimeUtc(path, longAgo.AddMinutes(minutes));
        return path;
      })
      .ToList();
    ReportFolder.Save(Reports, "the last", When.AddHours(1));

    var kept = Directory.GetFiles(Reports);
    Assert.Equal(ReportFolder.Kept, kept.Length);
    Assert.DoesNotContain(paths[0], kept);
    Assert.DoesNotContain(paths[2], kept);
    Assert.Contains(paths[3], kept);
  }

  [Fact]
  public void Other_files_in_the_folder_are_left_alone()
  {
    Directory.CreateDirectory(Reports);
    var other = Path.Combine(Reports, "notes.txt");
    File.WriteAllText(other, "mine");
    for (var minutes = 0; minutes <= ReportFolder.Kept; minutes++)
    {
      ReportFolder.Save(Reports, "a report", When.AddMinutes(minutes));
    }

    Assert.True(File.Exists(other));
  }
}
