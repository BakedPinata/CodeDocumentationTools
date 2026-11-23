using System.Text;
using CodeDocumenter.Workbooks;
using OpenMcdf;
using VbaDocumenter.Core.Extraction;
using VbaDocumenter.Core.Importing;

namespace S4Spike;

/// <summary>A form to add to a project: its code as the project stores it, its VBFrame text, and its designer storage.</summary>
sealed record FormToAdd(string Name, string Code, string VbFrame, Action<Storage> FillDesigner);

/// <summary>How the MSForms library is referred to.</summary>
enum FormsReference { None, Fm20, CopiedFromWorkbook }

/// <summary>One workbook to build, and what it tries.</summary>
sealed record Variant(string File, string Tries, Func<Builder, IReadOnlyList<FormToAdd>> Forms, FormsReference Reference);

/// <summary>
/// Builds new workbooks holding the sample's two forms, each a different way, with a module SpikeTest whose Check
/// function loads both and says what it found, and whose ShowForms shows them.
/// </summary>
sealed class Builder
{
  const string PackageLine = "Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}";
  const string Fm20 = @"*\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#C:\Windows\System32\FM20.DLL#Microsoft Forms 2.0 Object Library";
  static readonly Encoding Ansi = CodePages.Get(1252);

  public static readonly string[] FormNames = ["frmDemoForm1", "frmDemoForm2"];

  public static readonly IReadOnlyList<Variant> Variants =
  [
    new("F1-files.xlsm", "From the .frm and .frx: designer copied from the .frx, VBFrame from the Begin block, VB_Base with new GUIDs, a reference to FM20.DLL",
      b => b.FromFiles(vbBase: true, asWorkbookHasIt: false), FormsReference.Fm20),
    new("F2-no-vb-base.xlsm", "As F1, without VB_Base",
      b => b.FromFiles(vbBase: false, asWorkbookHasIt: false), FormsReference.Fm20),
    new("F3-no-reference.xlsm", "As F1, with no reference to the MSForms library",
      b => b.FromFiles(vbBase: true, asWorkbookHasIt: false), FormsReference.None),
    new("F4-designer-as-workbook.xlsm", "As F1, with the designer's class id and CompObj as a workbook has them (zeros, no ProgID)",
      b => b.FromFiles(vbBase: true, asWorkbookHasIt: true), FormsReference.Fm20),
    new("C1-copied.xlsm", "From the workbook: designer storage, VBFrame, code (with its VB_Base) and the MSForms reference copied as they are",
      b => b.FromWorkbook(), FormsReference.CopiedFromWorkbook),
  ];

  readonly string _sample;
  readonly string _workbook;

  public Builder(string sampleFolder)
  {
    _sample = sampleFolder;
    _workbook = Directory.GetFiles(sampleFolder, "*.xlsm").Single();
  }

  public static int Run(string sampleFolder, string outFolder, IReadOnlyCollection<string> only)
  {
    var builder = new Builder(sampleFolder);
    Directory.CreateDirectory(outFolder);
    foreach (var v in Variants.Where(v => only.Count == 0 || only.Any(o => v.File.StartsWith(o, StringComparison.OrdinalIgnoreCase))))
    {
      var bytes = builder.Build(v);
      File.WriteAllBytes(Path.Combine(outFolder, v.File), bytes);
      Console.WriteLine($"{v.File}: {v.Tries}");
    }

    return 0;
  }

  public byte[] Build(Variant v)
  {
    var source = TestModule.ReplaceLineEndings("\r\n");
    var test = Ansi.GetBytes(source);
    var module = new VbaModule { Name = "SpikeTest", Kind = VbaModuleKind.Standard, StreamName = "SpikeTest", SourceBytes = test, Source = source };
    var workbook = TargetWorkbook.New().Write(new ModuleChange([module], [], [], 1252));

    var package = OpcPackage.Read(workbook);
    var part = package.Names.First(n => n.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase));
    var references = v.Reference switch
    {
      FormsReference.Fm20 => Registered("MSForms", Fm20),
      FormsReference.CopiedFromWorkbook => DirRecords.Read(WorkbookReader.Read(File.ReadAllBytes(_workbook)).References.Single(r => r.Name == "MSForms").Records),
      _ => [],
    };
    var forms = v.Forms(this);
    package.Set(part, AddForms(package.Get(part), forms, references));
    var output = package.Write();

    // Read back: each form's code has to be there as written.
    var read = WorkbookReader.Read(output).Modules.ToDictionary(m => m.Name);
    foreach (var form in forms)
      if (!read.TryGetValue(form.Name, out var m) || m.Source != form.Code || m.Kind != VbaModuleKind.Form)
        throw new InvalidDataException($"{v.File}: {form.Name} didn't read back as written.");
    return output;
  }

  // From the VBA editor's files: the code after the Begin block, the Begin block as VBFrame, and the .frx's compound file.
  IReadOnlyList<FormToAdd> FromFiles(bool vbBase, bool asWorkbookHasIt) => [.. FormNames.Select(name =>
  {
    var frm = Ansi.GetString(File.ReadAllBytes(Path.Combine(_sample, name + ".frm"))).ReplaceLineEndings("\r\n");
    var lines = frm.Split("\r\n").ToList();
    int end = lines.FindIndex(l => l == "End");
    var frame = string.Join("\r\n", lines.Take(end + 1).Where(l => !l.TrimStart().StartsWith("OleObjectBlob", StringComparison.Ordinal))) + "\r\n";
    var code = string.Join("\r\n", lines.Skip(end + 1));
    code = FormAttributes(code, vbBase);
    var frx = File.ReadAllBytes(Path.Combine(_sample, name + ".frx"))[24..];
    return new FormToAdd(name, code, frame, designer =>
    {
      using var source = RootStorage.Open(new MemoryStream(frx));
      if (!asWorkbookHasIt) designer.CLSID = source.CLSID;
      CopyChildren(source, designer);
      if (asWorkbookHasIt)
      {
        // The workbook's CompObj: no class id and no ProgID.
        var compObj = Streams.Read(source, "\u0001CompObj");
        WriteStream(designer, "\u0001CompObj", WorkbookCompObj(compObj));
      }
    });
  })];

  // From the sample workbook, as it stores each form.
  IReadOnlyList<FormToAdd> FromWorkbook()
  {
    var bytes = File.ReadAllBytes(_workbook);
    var modules = WorkbookReader.Read(bytes).Modules.ToDictionary(m => m.Name);
    var package = OpcPackage.Read(bytes);
    var project = package.Get(package.Names.First(n => n.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase)));
    return [.. FormNames.Select(name =>
    {
      using var root = RootStorage.Open(new MemoryStream(project));
      var frame = Ansi.GetString(Streams.Read(root.OpenStorage(name), "\u0003VBFrame"));
      return new FormToAdd(name, modules[name].Source, frame, designer =>
      {
        using var source = RootStorage.Open(new MemoryStream(project));
        var storage = source.OpenStorage(name);
        designer.CLSID = storage.CLSID;
        CopyChildren(storage, designer, skip: "\u0003VBFrame");
      });
    })];
  }

  // The attributes a project stores for a form, as the VBA editor's import gives them: VB_Base names two GUIDs it makes.
  static string FormAttributes(string code, bool vbBase)
  {
    var lines = code.Split("\r\n").ToList();
    int name = lines.FindIndex(l => l.StartsWith("Attribute VB_Name", StringComparison.Ordinal));
    if (vbBase) lines.Insert(name + 1, $"Attribute VB_Base = \"0{{{Guid.NewGuid().ToString().ToUpperInvariant()}}}{{{Guid.NewGuid().ToString().ToUpperInvariant()}}}\"");
    int last = lines.FindLastIndex(l => l.StartsWith("Attribute VB_", StringComparison.Ordinal));
    lines.InsertRange(last + 1, ["Attribute VB_TemplateDerived = False", "Attribute VB_Customizable = False"]);
    return string.Join("\r\n", lines);
  }

  // A CompObj stream (MS-OLEDS 2.3.7) with its class id zeroed and no ProgID, which is how the workbook keeps a form's.
  static byte[] WorkbookCompObj(byte[] compObj)
  {
    var output = compObj.ToArray();
    Array.Clear(output, 12, 16);
    int at = 28;
    at += 4 + BitConverter.ToInt32(output, at);         // AnsiUserType
    at += 4 + BitConverter.ToInt32(output, at);         // AnsiClipboardFormat (a length-prefixed string here)
    int progId = BitConverter.ToInt32(output, at);       // Reserved1: the ProgID
    return [.. output.AsSpan(0, at), 0, 0, 0, 0, .. output.AsSpan(at + 4 + progId)];
  }

  static void CopyChildren(Storage from, Storage to, string? skip = null)
  {
    foreach (var entry in from.EnumerateEntries())
    {
      if (entry.Name == skip) continue;
      if (entry.Type == EntryType.Storage)
      {
        var child = to.CreateStorage(entry.Name);
        var source = from.OpenStorage(entry.Name);
        child.CLSID = source.CLSID;
        CopyChildren(source, child);
      }
      else WriteStream(to, entry.Name, Streams.Read(from, entry.Name));
    }
  }

  static void WriteStream(Storage storage, string name, byte[] data)
  {
    if (storage.ContainsEntry(name)) storage.Delete(name);
    using var stream = storage.CreateStream(name);
    stream.Write(data);
  }

  static List<DirRecord> Registered(string name, string libid)
  {
    var id = Ansi.GetBytes(libid);
    return
    [
      new(0x0016, Ansi.GetBytes(name)),
      new(0x003E, Encoding.Unicode.GetBytes(name)),
      new(0x000D, [.. BitConverter.GetBytes(id.Length), .. id, 0, 0, 0, 0, 0, 0]),
    ];
  }

  // Adds the forms to a project vbadoc wrote: code streams, dir records, PROJECT and PROJECTwm, and a designer storage each.
  static byte[] AddForms(byte[] vbaProjectBin, IReadOnlyList<FormToAdd> forms, List<DirRecord> references)
  {
    var buffer = new MemoryStream();
    buffer.Write(vbaProjectBin);
    buffer.Position = 0;
    using (var root = RootStorage.Open(buffer, StorageModeFlags.LeaveOpen))
    {
      var vba = root.OpenStorage("VBA");
      var dir = DirRecords.Read(OvbaDecompressor.Decompress(Streams.Read(vba, "dir")));
      var project = Ansi.GetString(Streams.Read(root, "PROJECT")).Split("\r\n").ToList();
      var nameMap = Streams.Read(root, "PROJECTwm").ToList();
      nameMap.RemoveRange(nameMap.Count - 2, 2); // the list's end

      foreach (var form in forms)
      {
        WriteStream(vba, form.Name, OvbaCompressor.Compress(Ansi.GetBytes(form.Code)));
        dir.InsertRange(dir.FindIndex(r => r.Id == 0x0010),
        [
          new(0x0019, Ansi.GetBytes(form.Name)),
          new(0x0047, Encoding.Unicode.GetBytes(form.Name)),
          new(0x001A, Ansi.GetBytes(form.Name)),
          new(0x0032, Encoding.Unicode.GetBytes(form.Name)),
          new(0x001C, []), new(0x0048, []),
          new(0x0031, [0, 0, 0, 0]),
          new(0x001E, [0, 0, 0, 0]),
          new(0x002C, [0xFF, 0xFF]),
          new(0x0022, []),
          new(0x0028, []),       // MODULEPRIVATE, as the sample has for each form
          new(0x002B, []),
        ]);

        int at = project.FindIndex(l => l.StartsWith("Name=", StringComparison.Ordinal));
        if (!project.Contains(PackageLine)) project.Insert(at++, PackageLine);
        project.Insert(at, $"BaseClass={form.Name}");
        project.Insert(project.Count - 1, $"{form.Name}=0, 0, 0, 0, C");

        nameMap.AddRange([.. Ansi.GetBytes(form.Name), 0, .. Encoding.Unicode.GetBytes(form.Name), 0, 0]);

        var designer = root.CreateStorage(form.Name);
        form.FillDesigner(designer);
        WriteStream(designer, "\u0003VBFrame", Ansi.GetBytes(form.VbFrame));
      }

      dir.InsertRange(dir.FindIndex(r => r.Id == 0x000F), references);
      dir.First(r => r.Id == 0x000F).Data = BitConverter.GetBytes((ushort)dir.Count(r => r.Id == 0x0019));
      WriteStream(vba, "dir", OvbaCompressor.Compress(DirRecords.Write(dir)));
      WriteStream(root, "PROJECT", Ansi.GetBytes(string.Join("\r\n", project)));
      WriteStream(root, "PROJECTwm", [.. nameMap, 0, 0]);
    }

    return buffer.ToArray();
  }

  const string TestModule = """
    Attribute VB_Name = "SpikeTest"
    Option Explicit

    ' Loads each form and says what it found, without showing either. Run from the Immediate window: ?Check
    Public Function Check() As String
        Dim s As String
        Load frmDemoForm1
        s = "frmDemoForm1: " & frmDemoForm1.Caption & ", " & frmDemoForm1.Controls.Count & " controls (" & ControlNames(frmDemoForm1) & "), list of " & frmDemoForm1.lbxDemoListBox.ListCount & " rows"
        Unload frmDemoForm1
        Load frmDemoForm2
        s = s & vbLf & "frmDemoForm2: " & frmDemoForm2.Caption & ", " & frmDemoForm2.Controls.Count & " controls (" & ControlNames(frmDemoForm2) & ")"
        Unload frmDemoForm2
        Check = s
    End Function

    Private Function ControlNames(ByVal form As Object) As String
        Dim c As Object
        For Each c In form.Controls
            ControlNames = ControlNames & c.Name & " "
        Next
        ControlNames = Trim$(ControlNames)
    End Function

    ' Shows each form in turn. Close one to see the next.
    Public Sub ShowForms()
        frmDemoForm1.Show
        frmDemoForm2.Show
    End Sub

    """;
}
