using CodeDocumenter.Shared;
using CodeDocumenter.Workbooks;
using CodeDocumenter.Workbooks.Importing;
using OpenMcdf;
using Serilog.Events;
using VbaDocumenter.Cli;
using VbaDocumenter.Cli.Logging;
using VbaDocumenter.Core.Extraction;
using VbaDocumenter.Core.Importing;
using VbaDocumenter.Tests.Support;

namespace VbaDocumenter.Tests.Cli;

/// <summary>
/// <c>vbadoc import</c> taking UserForms (step S4b): from a folder of the VBA editor's <c>.frm</c> and <c>.frx</c> files,
/// or from a workbook's designer, into a new workbook, a copy of a target and a bundle. What Excel accepted is in step S4a.
/// </summary>
public sealed class ImportFormsTests : IDisposable
{
  static readonly string FormsFixture = Path.Combine(AppContext.BaseDirectory, "Fixtures", "UserForms");
  static readonly string[] FormFiles = ["frmDemoForm1.frm", "frmDemoForm1.frx", "frmDemoForm2.frm", "frmDemoForm2.frx"];

  readonly string _folder = Path.Combine(Path.GetTempPath(), $"vbadoc-forms-{Guid.NewGuid():N}");
  readonly TestLog _log = new();
  readonly StringWriter _output = new();
  readonly StringWriter _error = new();
  FakeUserInteraction _interaction = new(isInteractive: false);

  string Forms => Path.Combine(_folder, "Forms");
  string New => Path.Combine(_folder, "Toolkit.xlsm");
  string Target => Path.Combine(_folder, "Target.xlsm");
  string Copy => Path.Combine(_folder, "Target-imported.xlsm");
  string Bundle => Path.Combine(_folder, "Bundle");

  public ImportFormsTests()
  {
    Directory.CreateDirectory(Forms);
    foreach (var file in FormFiles) File.Copy(Path.Combine(FormsFixture, file), Path.Combine(Forms, file));
    File.WriteAllText(Path.Combine(Forms, "Launch.bas"), "Attribute VB_Name = \"Launch\"\r\nOption Explicit\r\nPublic Sub Go()\r\n    frmDemoForm1.Show\r\nEnd Sub\r\n");
  }

  public void Dispose()
  {
    _output.Dispose();
    _error.Dispose();
    Directory.Delete(_folder, recursive: true);
  }

  int Run(string source, params string[] args) => CliRunner.Run(["import", source, .. args], new AppServices
  {
    Output = _output,
    Error = _error,
    Interaction = _interaction,
    IsLibraryInstalled = _ => true,
    CreateLogger = verbosity => _log.CreateLogger(LoggingSetup.MinimumLevel(verbosity)),
  }, TestContext.Current.CancellationToken);

  static Dictionary<string, VbaModule> Modules(string path) => WorkbookReader.Read(path).Modules.ToDictionary(m => m.Name);

  static VbaModule FromFolder(string folder, string name) => ModuleFolderReader.Read(folder, recursive: false, folderGroups: false).Project.Modules.Single(m => m.Name == name);

  static string ProjectText(string workbook)
  {
    var package = OpcPackage.Read(File.ReadAllBytes(workbook));
    using var root = RootStorage.Open(new MemoryStream(package.Get(package.Names.Single(n => n.EndsWith("vbaProject.bin", StringComparison.Ordinal)))));
    using var stream = root.OpenStream("PROJECT");
    return new StreamReader(stream, System.Text.Encoding.Latin1).ReadToEnd();
  }

  [Fact]
  public void Forms_from_the_VBA_editors_files_go_into_a_new_workbook_with_their_designers()
  {
    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "frmDemoForm1,frmDemoForm2", "--new", New));

    var modules = Modules(New);
    foreach (var name in new[] { "frmDemoForm1", "frmDemoForm2" })
    {
      var form = modules[name];
      Assert.Equal(VbaModuleKind.Form, form.Kind);
      Assert.True(form.Designer!.SameAs(FromFolder(Forms, name).Designer!), name);
      Assert.Matches(@"Attribute VB_Base = ""0\{[0-9A-F-]{36}\}\{[0-9A-F-]{36}\}""\r\n", form.Source);
      Assert.Contains("Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = False\r\n", form.Source, StringComparison.Ordinal);
      Assert.DoesNotContain("OleObjectBlob", form.Designer.Frame, StringComparison.Ordinal);
    }

    Assert.Contains("Private Sub UserForm_Initialize()", modules["frmDemoForm1"].Source, StringComparison.Ordinal);
    var project = ProjectText(New);
    Assert.Contains("Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}\r\nBaseClass=frmDemoForm1\r\nBaseClass=frmDemoForm2\r\n", project, StringComparison.Ordinal);
  }

  // A form's code that names an MSForms type needs the library, and a folder has no references to bring (step S4a).
  [Fact]
  public void Forms_bring_a_reference_to_the_MSForms_library_when_the_target_has_none()
  {
    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "frmDemoForm2", "--new", New));

    var reference = WorkbookReader.Read(New).References.Single(r => r.Name == "MSForms");
    Assert.Equal(FormsLibrary.Guid, reference.Guid);
    Assert.Equal(VbaReferenceKind.Registered, reference.Kind);
    Assert.Contains(_log.Messages(LogEventLevel.Information), m => m.Contains("Adding the library reference: Microsoft Forms 2.0 Object Library (MSForms)", StringComparison.Ordinal));
  }

  [Fact]
  public void A_module_that_shows_a_form_brings_it_along()
  {
    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "Launch", "--new", New));

    Assert.Contains("frmDemoForm1", Modules(New).Keys);
    Assert.DoesNotContain("frmDemoForm2", Modules(New).Keys);
  }

  [Fact]
  public void Forms_from_a_workbook_are_copied_with_their_designers_and_the_workbooks_reference()
  {
    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "frmDemoForm1,frmDemoForm2", "--new", Target));
    var source = Modules(Target);

    Assert.Equal(ExitCodes.Success, Run(Target, "--items", "frmDemoForm1,frmDemoForm2", "--new", New));

    var copied = Modules(New);
    foreach (var name in new[] { "frmDemoForm1", "frmDemoForm2" })
    {
      Assert.Equal(source[name].Source, copied[name].Source);
      Assert.True(copied[name].Designer!.SameAs(source[name].Designer!), name);
    }

    Assert.Single(WorkbookReader.Read(New).References, r => r.Name == "MSForms");
  }

  [Fact]
  public void Forms_go_into_a_copy_of_a_target_that_has_a_form_already_beside_it()
  {
    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "frmDemoForm2", "--new", Target));
    var before = File.ReadAllBytes(Target);

    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "frmDemoForm1", "--into", Target));

    var modules = Modules(Copy);
    Assert.True(modules["frmDemoForm1"].Designer!.SameAs(FromFolder(Forms, "frmDemoForm1").Designer!));
    Assert.True(modules["frmDemoForm2"].Designer!.SameAs(FromFolder(Forms, "frmDemoForm2").Designer!));
    var project = ProjectText(Copy);
    Assert.Single(project.Split("\r\n"), l => l.StartsWith("Package=", StringComparison.Ordinal));
    Assert.Single(WorkbookReader.Read(Copy).References, r => r.Name == "MSForms");
    Assert.Equal(before, File.ReadAllBytes(Target));
  }

  [Fact]
  public void A_form_the_target_has_is_left_alone_as_identical_when_it_is_the_same()
  {
    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "frmDemoForm1", "--new", Target));

    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "frmDemoForm1", "--into", Target));

    Assert.Contains(_log.Messages(LogEventLevel.Information), m => m == "Already there, identical: frmDemoForm1");
    Assert.False(File.Exists(Copy));
  }

  [Fact]
  public void A_different_form_the_target_has_can_only_be_skipped()
  {
    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "frmDemoForm1", "--new", Target));
    var frm = Path.Combine(Forms, "frmDemoForm1.frm");
    File.WriteAllText(frm, File.ReadAllText(frm).Replace("Private Sub UserForm_Initialize()", "Private Sub UserForm_Activate()", StringComparison.Ordinal));

    Assert.Equal(ExitCodes.OutputRefused, Run(Forms, "--items", "frmDemoForm1", "--into", Target, "--on-clash", "overwrite"));
    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "frmDemoForm1,frmDemoForm2", "--into", Target, "--on-clash", "skip"));

    Assert.Contains("Private Sub UserForm_Initialize()", Modules(Copy)["frmDemoForm1"].Source, StringComparison.Ordinal);
    Assert.Contains("frmDemoForm2", Modules(Copy).Keys);
  }

  [Theory]
  [InlineData("frx", "Its .frx, frmDemoForm1.frx, isn't beside it")]
  [InlineData("begin", "Its .frm has no designer (no Begin block)")]
  public void A_form_without_its_designer_cant_be_taken_and_says_why(string missing, string why)
  {
    var frm = Path.Combine(Forms, "frmDemoForm1.frm");
    if (missing == "frx") File.Delete(Path.Combine(Forms, "frmDemoForm1.frx"));
    else File.WriteAllText(frm, File.ReadAllText(frm)[File.ReadAllText(frm).IndexOf("Attribute VB_Name", StringComparison.Ordinal)..]);

    Assert.Equal(ExitCodes.InputError, Run(Forms, "--items", "frmDemoForm1", "--new", New));

    Assert.Contains(_log.Messages(LogEventLevel.Error), m => m.StartsWith("frmDemoForm1 can't be taken: " + why, StringComparison.Ordinal));
    Assert.False(File.Exists(New));
    Assert.Contains(ImportSource.Items(Forms, recursive: false), i => i.Name == "frmDemoForm1" && i.WhyNot!.StartsWith(why, StringComparison.Ordinal));
  }

  // The VBA editor's own files come out as they went in: the .frx as it was, and the .frm with its Begin block.
  [Fact]
  public void A_bundle_from_a_folder_has_each_form_as_the_VBA_editor_exported_it()
  {
    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "frmDemoForm1,frmDemoForm2", "--bundle", Bundle));

    foreach (var file in FormFiles) Assert.Equal(File.ReadAllBytes(Path.Combine(Forms, file)), File.ReadAllBytes(Path.Combine(Bundle, file)));
    Assert.Contains("keep the two together", File.ReadAllText(Path.Combine(Bundle, "README.md")), StringComparison.Ordinal);
  }

  [Fact]
  public void A_bundle_from_a_workbook_has_each_forms_frm_as_exported_and_a_frx_that_reads_back()
  {
    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "frmDemoForm1,frmDemoForm2", "--new", Target));

    Assert.Equal(ExitCodes.Success, Run(Target, "--items", "frmDemoForm1,frmDemoForm2", "--bundle", Bundle));

    foreach (var name in new[] { "frmDemoForm1", "frmDemoForm2" })
    {
      Assert.Equal(File.ReadAllText(Path.Combine(Forms, name + ".frm")), File.ReadAllText(Path.Combine(Bundle, name + ".frm")));
      // The header gives the size of the compound file after it, which vbadoc lays out its own way, and otherwise has what
      // the VBA editor wrote.
      var frx = File.ReadAllBytes(Path.Combine(Bundle, name + ".frx"));
      var original = File.ReadAllBytes(Path.Combine(Forms, name + ".frx"));
      Assert.Equal(frx.Length - 24, BitConverter.ToInt32(frx, 4));
      Assert.Equal([.. original[..4], .. original[8..24]], [.. frx[..4], .. frx[8..24]]);
      Assert.True(FromFolder(Bundle, name).Designer!.SameAs(FromFolder(Forms, name).Designer!), name);
    }
  }

  [Fact]
  public void An_earlier_bundles_frx_goes_when_its_form_does()
  {
    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "frmDemoForm1", "--bundle", Bundle));

    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "Launch", "--bundle", Bundle, "--force"));
    Assert.Equal(ExitCodes.Success, Run(Forms, "--items", "frmDemoForm2", "--bundle", Bundle, "--force"));

    Assert.False(File.Exists(Path.Combine(Bundle, "frmDemoForm1.frx")));
    Assert.True(File.Exists(Path.Combine(Bundle, "frmDemoForm2.frx")));
  }
}
