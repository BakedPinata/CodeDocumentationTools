using CodeDocumenter.Shared;
using Serilog.Events;
using VbaDocumenter.Cli;
using VbaDocumenter.Cli.Logging;
using VbaDocumenter.Core.Extraction;
using VbaDocumenter.Tests.Support;

namespace VbaDocumenter.Tests.Cli;

/// <summary>
/// vbadoc's export writes each UserForm as the VBA editor does (step S4c): a <c>.frm</c> with the designer's
/// <c>Begin</c> block and a <c>.frx</c> beside it, so its own folders bring forms back, with <c>import</c> or by hand.
/// </summary>
public sealed class ExportFormsTests : IDisposable
{
  static readonly string FormsFixture = Path.Combine(AppContext.BaseDirectory, "Fixtures", "UserForms");
  static readonly string[] Forms = ["frmDemoForm1", "frmDemoForm2"];

  readonly string _folder = Path.Combine(Path.GetTempPath(), $"vbadoc-export-forms-{Guid.NewGuid():N}");
  readonly TestLog _log = new();
  readonly StringWriter _output = new();
  readonly StringWriter _error = new();

  string Workbook => Path.Combine(_folder, "Forms.xlsm");
  string Exported => Path.Combine(_folder, "Exported");

  public ExportFormsTests()
  {
    Directory.CreateDirectory(_folder);
    Assert.Equal(ExitCodes.Success, Run("import", FormsFixture, "--items", "frmDemoForm1,frmDemoForm2", "--new", Workbook));
  }

  public void Dispose()
  {
    _output.Dispose();
    _error.Dispose();
    Directory.Delete(_folder, recursive: true);
  }

  int Run(params string[] args) => CliRunner.Run(args, new AppServices
  {
    Output = _output,
    Error = _error,
    Interaction = new FakeUserInteraction(isInteractive: false),
    IsLibraryInstalled = _ => true,
    CreateLogger = verbosity => _log.CreateLogger(LoggingSetup.MinimumLevel(verbosity)),
  }, TestContext.Current.CancellationToken);

  static FormDesigner DesignerIn(string folder, string form) =>
    ModuleFolderReader.Read(folder, recursive: false, folderGroups: false).Project.Modules.Single(m => m.Name == form).Designer!;

  [Fact]
  public void Each_form_is_written_as_the_VBA_editor_exported_it()
  {
    Assert.Equal(ExitCodes.Success, Run(Workbook, "--output", Exported, "--force"));

    foreach (var form in Forms)
    {
      Assert.Equal(File.ReadAllBytes(Path.Combine(FormsFixture, form + ".frm")), File.ReadAllBytes(Path.Combine(Exported, form + ".frm")));
      Assert.True(DesignerIn(Exported, form).SameAs(DesignerIn(FormsFixture, form)), form);
      Assert.Equal(File.ReadAllBytes(Path.Combine(FormsFixture, form + ".frx"))[8..24], File.ReadAllBytes(Path.Combine(Exported, form + ".frx"))[8..24]);
    }
  }

  // The folder is kept in git, so the same workbook has to give the same files.
  [Fact]
  public void The_same_workbook_gives_the_same_files_each_time()
  {
    Assert.Equal(ExitCodes.Success, Run(Workbook, "--output", Exported, "--force"));
    var first = Directory.GetFiles(Exported).ToDictionary(f => Path.GetFileName(f), File.ReadAllBytes);

    Assert.Equal(ExitCodes.Success, Run(Workbook, "--output", Exported, "--force"));

    Assert.Contains(_log.Messages(LogEventLevel.Information), m => m.Contains("0 written", StringComparison.Ordinal));
    foreach (var (name, bytes) in first) Assert.Equal(bytes, File.ReadAllBytes(Path.Combine(Exported, name)));
  }

  [Fact]
  public void Its_own_folder_brings_the_forms_back()
  {
    Assert.Equal(ExitCodes.Success, Run(Workbook, "--output", Exported, "--force"));
    var again = Path.Combine(_folder, "Again.xlsm");

    Assert.Equal(ExitCodes.Success, Run("import", Exported, "--items", "frmDemoForm1,frmDemoForm2", "--new", again));

    var modules = WorkbookReader.Read(again).Modules.ToDictionary(m => m.Name);
    foreach (var form in Forms) Assert.True(modules[form].Designer!.SameAs(DesignerIn(FormsFixture, form)), form);
  }

  [Fact]
  public void A_frx_goes_when_its_form_does()
  {
    Assert.Equal(ExitCodes.Success, Run(Workbook, "--output", Exported, "--force"));
    var fewer = Path.Combine(_folder, "Fewer.xlsm");
    Assert.Equal(ExitCodes.Success, Run("import", FormsFixture, "--items", "frmDemoForm2", "--new", fewer));

    Assert.Equal(ExitCodes.Success, Run(fewer, "--output", Exported, "--force"));

    Assert.False(File.Exists(Path.Combine(Exported, "frmDemoForm1.frx")));
    Assert.True(File.Exists(Path.Combine(Exported, "frmDemoForm2.frx")));
  }

  // The page shows the file as it's written, so a form's source starts with its Begin block.
  [Fact]
  public void The_page_has_each_form_as_its_file()
  {
    Assert.Equal(ExitCodes.Success, Run(Workbook, "--output", Exported, "--force", "--embed-source"));

    var page = File.ReadAllText(Path.Combine(Exported, "Forms-vba.html"));
    Assert.Contains("OleObjectBlob   =   &quot;frmDemoForm1.frx&quot;:0000", page, StringComparison.Ordinal);
  }

  // A form whose designer can't be read is still exported, as its code, with a warning (Requirements, Forms).
  [Fact]
  public void A_form_without_its_designer_is_written_as_its_code_with_a_warning()
  {
    var synthetic = Path.Combine(_folder, "Synthetic.xlsm");
    File.WriteAllBytes(synthetic, SyntheticWorkbook.Xlsm([new("Lonely", VbaModuleKind.Form, "Attribute VB_Name = \"Lonely\"\r\nOption Explicit\r\n")]));

    Assert.Equal(ExitCodes.Success, Run(synthetic, "--output", Exported, "--force"));

    Assert.Equal("Attribute VB_Name = \"Lonely\"\r\nOption Explicit\r\n", File.ReadAllText(Path.Combine(Exported, "Lonely.frm")));
    Assert.False(File.Exists(Path.Combine(Exported, "Lonely.frx")));
    Assert.Contains(_log.Messages(LogEventLevel.Warning), m => m.StartsWith("Lonely is written as its code alone", StringComparison.Ordinal));
  }
}
