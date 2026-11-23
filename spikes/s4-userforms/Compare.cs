using System.Text;
using CodeDocumenter.Workbooks;
using OpenMcdf;
using VbaDocumenter.Core.Extraction;

namespace S4Spike;

/// <summary>Compares a form's storage in a workbook with its .frx, stream by stream, and prints the project's references.</summary>
static class Compare
{
  public static int Run(string workbook, string frx)
  {
    var bytes = File.ReadAllBytes(workbook);
    var name = Path.GetFileNameWithoutExtension(frx);
    var package = OpcPackage.Read(bytes);
    var part = package.Names.First(n => n.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase));
    using var project = RootStorage.Open(new MemoryStream(package.Get(part)));
    using var file = RootStorage.Open(new MemoryStream(File.ReadAllBytes(frx)[24..]));
    Console.WriteLine($"frx root clsid {file.CLSID}; workbook storage clsid {project.OpenStorage(name).CLSID}");
    Walk(project.OpenStorage(name), file, "");

    foreach (var reference in WorkbookReader.Read(bytes).References)
      Console.WriteLine($"reference {reference.Name} {reference.Kind} {reference.Guid} {reference.Libid}");
    return 0;
  }

  static void Walk(Storage a, Storage b, string path)
  {
    var names = a.EnumerateEntries().Select(e => e.Name).Union(b.EnumerateEntries().Select(e => e.Name));
    foreach (var n in names)
    {
      var p = path + "/" + Inspector.Printable(n);
      if (!a.ContainsEntry(n) || !b.ContainsEntry(n)) { Console.WriteLine($"{p}: only in {(a.ContainsEntry(n) ? "workbook" : "frx")}"); continue; }
      if (a.TryOpenStorage(n, out var sa) && b.TryOpenStorage(n, out var sb)) { Walk(sa!, sb!, p); continue; }
      var x = Streams.Read(a, n); var y = Streams.Read(b, n);
      if (x.AsSpan().SequenceEqual(y)) { Console.WriteLine($"{p}: same ({x.Length})"); continue; }
      Console.WriteLine($"{p}: DIFFERENT");
      Console.WriteLine($"  workbook {Convert.ToHexString(x)}\n  {Encoding.Latin1.GetString(x).Replace('\0', '.')}");
      Console.WriteLine($"  frx      {Convert.ToHexString(y)}\n  {Encoding.Latin1.GetString(y).Replace('\0', '.')}");
    }
  }
}
