using VbaDocumenter.Core.Extraction;

namespace AccessSpike;

/// <summary>A module's name and source, however it was got.</summary>
public sealed record ModuleText(string Name, string Text);

/// <summary>What choice C found in a database file.</summary>
/// <param name="Modules">The VBA project's modules, with their source.</param>
/// <param name="Storage">Every storage and stream in the file's own store, as a path and a size.</param>
public sealed record DirectResult(IReadOnlyList<ModuleText> Modules, IReadOnlyList<(string Path, int Size)> Storage);

/// <summary>
/// Choice C: the VBA read from the database file with no Access. A database keeps its VBA project, and each form's
/// design, in a system table, MSysAccessStorage, which is a small file system: each row is a storage (a folder) or a
/// stream (a file), with its parent's id. The VBA project's streams are the ones a workbook's vbaProject.bin has
/// ([MS-OVBA]), so vbadoc's reading of them does the rest.
/// </summary>
public static class DirectReader
{
  public static DirectResult Read(string databasePath)
  {
    var file = new JetFile(File.ReadAllBytes(databasePath));
    if (!file.Tables().TryGetValue("MSysAccessStorage", out var page))
    {
      throw new InvalidDataException("The file has no MSysAccessStorage table.");
    }

    var rows = file.Rows(file.ReadTable(page)).Where(r => r["Id"] is int).ToList();
    var byId = rows.ToDictionary(r => (int)r["Id"]!);
    string PathOf(Dictionary<string, object?> row)
    {
      var names = new List<string>();
      for (var at = row; at is not null && names.Count < 32;)
      {
        names.Add((string?)at["Name"] ?? "?");
        at = at["ParentId"] is int parent && parent != (int)at["Id"]! && byId.TryGetValue(parent, out var up) ? up : null;
      }

      names.Reverse();
      return string.Join("/", names);
    }

    var storage = rows.Select(r => (Path: PathOf(r), Size: r["Lv"] is byte[] value ? value.Length : -1)).OrderBy(s => s.Path, StringComparer.Ordinal).ToList();

    // The VBA project's own streams: "dir" lists the modules, and each module's stream is beside it.
    var modules = new List<ModuleText>();
    var dir = rows.FirstOrDefault(r => (string?)r["Name"] == "dir" && r["Lv"] is byte[] && PathOf(r).Contains("/VBA/", StringComparison.Ordinal));
    if (dir is not null)
    {
      var info = DirStreamParser.Parse(OvbaDecompressor.Decompress((byte[])dir["Lv"]!));
      var encoding = CodePages.Get(info.CodePage);
      var beside = rows.Where(r => Equals(r["ParentId"], dir["ParentId"]) && r["Lv"] is byte[])
        .ToDictionary(r => (string)r["Name"]!, r => (byte[])r["Lv"]!, StringComparer.OrdinalIgnoreCase);
      foreach (var module in info.Modules)
      {
        if (beside.TryGetValue(module.StreamName, out var stream))
        {
          modules.Add(new ModuleText(module.Name, encoding.GetString(OvbaDecompressor.Decompress(stream, (int)module.TextOffset))));
        }
      }
    }

    return new DirectResult(modules, storage);
  }
}
