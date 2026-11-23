using System.Text;
using CodeDocumenter.Workbooks;
using OpenMcdf;
using VbaDocumenter.Core.Extraction;
using VbaDocumenter.Core.Importing;

namespace S4Spike;

/// <summary>Prints every record of a workbook's dir stream.</summary>
static class DirDump
{
  public static int Run(string workbook)
  {
    var package = OpcPackage.Read(File.ReadAllBytes(workbook));
    var part = package.Names.First(n => n.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase));
    using var root = RootStorage.Open(new MemoryStream(package.Get(part)));
    foreach (var r in DirRecords.Read(OvbaDecompressor.Decompress(Streams.Read(root.OpenStorage("VBA"), "dir"))))
      Console.WriteLine($"{r.Id:X4} {r.Data.Length,4} {Convert.ToHexString(r.Data.AsSpan(0, Math.Min(r.Data.Length, 48)))}  {Encoding.Latin1.GetString(r.Data.AsSpan(0, Math.Min(r.Data.Length, 48))).Replace('\0', '.')}");
    return 0;
  }
}
