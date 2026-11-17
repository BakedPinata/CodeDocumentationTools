using CodeDocumenter.Desktop.Documenting;
using CodeDocumenter.Desktop.Errors;
using CodeDocumenter.Desktop.Importing;
using CodeDocumenter.Desktop.Running;

namespace CodeDocumenter.Desktop.Tests.Errors;

public sealed class StateSummaryTests : IDisposable
{
  private readonly TempFolder folder = new();

  public void Dispose() => folder.Dispose();

  [Fact]
  public void A_path_is_described_by_its_kind_only()
  {
    var file = folder.Combine("Payroll.xlsm");
    File.WriteAllText(file, "");

    Assert.Equal("none", StateSummary.KindOf("  "));
    Assert.Equal("a folder", StateSummary.KindOf(folder.Path));
    Assert.Equal("a file", StateSummary.KindOf($"\"{file}\""));
    Assert.Equal("not there", StateSummary.KindOf(folder.Combine("Gone.xlsm")));
  }

  [Fact]
  public void What_the_app_was_set_to_is_in_kinds_and_counts_with_no_path_or_name()
  {
    var modules = folder.Combine("Payroll modules");
    Directory.CreateDirectory(modules);
    var state = new AppState(
      "Import",
      Running: false,
      ScalePercent: 150,
      new DocumentSettings { Input = folder.Combine("Payroll.xlsm"), Pqdoc = new PqdocSettings { Run = false } },
      new ImportSettings
      {
        Source = modules,
        Recursive = true,
        Items = ["PayrollCalc", "SalaryTable"],
        WriteTo = ImportDestination.NewWorkbook,
        NewWorkbook = folder.Combine("Payroll new.xlsm"),
        Bundle = folder.Combine("Payroll bundle"),
        Target = folder.Combine("Payroll target.xlsm"),
        Manifest = folder.Combine("Payroll manifest.json"),
      },
      Listed: 40,
      Recent: 3);

    var lines = StateSummary.Describe(state);

    Assert.Equal(
      new[]
      {
        "Tab showing: Import",
        "A run going: no",
        "Screen scale: 150%",
        "Document tab: input not there, subfolders no; vbadoc ticked; pqdoc not ticked, 1 formats",
        "Import tab: source a folder, subfolders yes; vbadoc; choosing from the list, 40 listed and 2 ticked; writing to a new workbook",
        "Recent workbooks and folders: 3",
      },
      lines);
    Assert.DoesNotContain(lines, line => line.Contains("Payroll", StringComparison.OrdinalIgnoreCase)
      || line.Contains("Salary", StringComparison.OrdinalIgnoreCase)
      || line.Contains(folder.Path, StringComparison.OrdinalIgnoreCase));
  }

  [Fact]
  public void A_manifest_is_described_by_its_kind_too()
  {
    var state = new AppState(
      "Document",
      Running: true,
      ScalePercent: 100,
      new DocumentSettings(),
      new ImportSettings { ChooseFrom = ChooseFrom.Manifest, Tool = Tool.Pqdoc, WriteTo = ImportDestination.InPlace },
      Listed: 0,
      Recent: 0);

    var import = Assert.Single(StateSummary.Describe(state), line => line.StartsWith("Import tab:", StringComparison.Ordinal));

    Assert.Equal("Import tab: source none, subfolders no; pqdoc; choosing from a manifest (none); writing to the target itself", import);
  }
}
