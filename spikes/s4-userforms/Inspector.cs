using System.Text;
using CodeDocumenter.Workbooks;
using OpenMcdf;

namespace S4Spike;

/// <summary>Prints what's inside a workbook's VBA project, or a .frx, so the two can be compared.</summary>
static class Inspector
{
  public static int Run(string path)
  {
    var bytes = File.ReadAllBytes(path);
    if (path.EndsWith(".frx", StringComparison.OrdinalIgnoreCase))
    {
      Console.WriteLine($"header: {Convert.ToHexString(bytes, 0, 24)}");
      using var root = RootStorage.Open(new MemoryStream(bytes[24..]));
      Tree(root, "");
      return 0;
    }

    var package = OpcPackage.Read(bytes);
    var part = package.Names.First(n => n.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase));
    using (var root = RootStorage.Open(new MemoryStream(package.Get(part))))
    {
      Tree(root, "");
      Console.WriteLine("--- PROJECT");
      Console.WriteLine(Encoding.Latin1.GetString(Streams.Read(root, "PROJECT")));
      foreach (var storage in root.EnumerateEntries().Where(e => e.Type == EntryType.Storage && e.Name != "VBA"))
      {
        var form = root.OpenStorage(storage.Name);
        if (!form.ContainsEntry("\u0003VBFrame")) continue;
        Console.WriteLine($"--- {storage.Name}/VBFrame");
        Console.WriteLine(Encoding.Latin1.GetString(Streams.Read(form, "\u0003VBFrame")));
      }
    }

    foreach (var module in VbaDocumenter.Core.Extraction.WorkbookReader.Read(bytes).Modules)
    {
      Console.WriteLine($"--- module {module.Name} ({module.Kind}, stream {module.StreamName}), its attributes:");
      foreach (var line in module.Source.Split("\r\n").TakeWhile(l => l.StartsWith("Attribute ", StringComparison.Ordinal)))
        Console.WriteLine("  " + line);
    }

    return 0;
  }

  static void Tree(Storage storage, string indent)
  {
    foreach (var entry in storage.EnumerateEntries())
    {
      Console.WriteLine($"{indent}{Printable(entry.Name)} {(entry.Type == EntryType.Storage ? "/" : entry.Length + " bytes")} clsid {entry.CLSID}");
      if (entry.Type == EntryType.Storage) Tree(storage.OpenStorage(entry.Name), indent + "  ");
    }
  }

  public static string Printable(string name) => string.Concat(name.Select(c => c < ' ' ? $"\\x{(int)c:X2}" : c.ToString()));
}

static class Streams
{
  public static byte[] Read(Storage storage, string name)
  {
    using var stream = storage.OpenStream(name);
    var buffer = new MemoryStream();
    stream.CopyTo(buffer);
    return buffer.ToArray();
  }
}
