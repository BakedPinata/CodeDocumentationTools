using System.Buffers.Binary;
using System.Text;
using System.Text.RegularExpressions;
using OpenMcdf;

namespace VbaDocumenter.Core.Extraction;

/// <summary>
/// A UserForm's designer: what a workbook keeps in the storage named after the form, and the VBA editor's export splits
/// between the <c>.frm</c>'s <c>Begin</c> block and the <c>.frx</c>. Nothing in it is understood, only moved (step S4a,
/// DECISIONS): the form and its controls are streams in a format of their own, copied as they are.
/// </summary>
/// <param name="Frame">
/// The form's <c>VBFrame</c> text, from <c>VERSION 5.00</c> to its <c>End</c>, with CRLF line endings and no
/// <c>OleObjectBlob</c> line, as a workbook keeps it.
/// </param>
/// <param name="Storage">
/// A compound file whose root is the designer: its class id, its streams (<c>f</c>, <c>o</c>, <c>CompObj</c>) and a
/// storage for each control that holds others, without <c>VBFrame</c>. It's what a <c>.frx</c> holds after its header.
/// </param>
public sealed partial record FormDesigner(string Frame, byte[] Storage)
{
  /// <summary>The name of the stream a workbook keeps <see cref="Frame"/> in.</summary>
  public const string FrameStream = "\u0003VBFrame";

  /// <summary>The length of the header a <c>.frx</c> has before its compound file.</summary>
  public const int FrxHeaderLength = 24;

  /// <summary>The <c>.frx</c> it was read from, when it came from a folder, so a bundle can have it as it was.</summary>
  public byte[]? Frx { get; init; }

  /// <summary>The designer in a workbook's project: the storage named after the form.</summary>
  /// <param name="storage">The form's storage.</param>
  /// <param name="encoding">The project's code page, which <c>VBFrame</c> is in.</param>
  public static FormDesigner FromStorage(Storage storage, Encoding encoding)
  {
    ArgumentNullException.ThrowIfNull(storage);
    ArgumentNullException.ThrowIfNull(encoding);
    var frame = storage.ContainsEntry(FrameStream) ? encoding.GetString(Read(storage, FrameStream)) : "";
    var buffer = new MemoryStream();
    using (var root = RootStorage.Create(buffer, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen))
    {
      root.CLSID = storage.CLSID;
      CopyChildren(storage, root, skip: FrameStream);
      CopyTimes(storage, root);
    }

    return new FormDesigner(frame.ReplaceLineEndings("\r\n"), buffer.ToArray());
  }

  /// <summary>
  /// The designer in an exported <c>.frm</c> and its <c>.frx</c>, or why there isn't one: the <c>.frm</c> has no
  /// <c>Begin</c> block (vbadoc's own export before step S4c), or its <c>.frx</c> is missing or can't be read.
  /// </summary>
  /// <param name="frm">The <c>.frm</c>'s text.</param>
  /// <param name="readFrx">Reads the <c>.frx</c> it names, by its name as written, or returns null when there's none.</param>
  public static (FormDesigner? Designer, string? Problem) FromFiles(string frm, Func<string, byte[]?> readFrx)
  {
    ArgumentNullException.ThrowIfNull(frm);
    ArgumentNullException.ThrowIfNull(readFrx);
    var text = frm.ReplaceLineEndings("\r\n");
    if (BeginBlock().Match(text) is not { Success: true } block)
      return (null, "Its .frm has no designer (no Begin block), so the form's controls aren't there to take. Export it from the VBA editor to get them");

    var lines = block.Value.Split("\r\n");
    var blob = lines.Select(l => OleObjectBlob().Match(l)).FirstOrDefault(m => m.Success);
    if (blob is null) return (null, "Its .frm's Begin block doesn't name its .frx (no OleObjectBlob line)");

    var name = blob.Groups["file"].Value;
    var offset = Convert.ToInt32(blob.Groups["offset"].Value, 16);
    var frx = readFrx(name);
    if (frx is null) return (null, $"Its .frx, {name}, isn't beside it, and the form's controls are in that file");

    if (frx.Length < offset + FrxHeaderLength || frx[offset] != 'L' || frx[offset + 1] != 'B')
      return (null, $"{name} isn't a .frx the VBA editor wrote");
    var size = BinaryPrimitives.ReadInt32LittleEndian(frx.AsSpan(offset + 4));
    if (size <= 0 || offset + FrxHeaderLength + size > frx.Length) return (null, $"{name} is cut short");

    var storage = frx.AsSpan(offset + FrxHeaderLength, size).ToArray();
    try
    {
      using var root = RootStorage.Open(new MemoryStream(storage));
      if (!root.ContainsEntry("f")) return (null, $"{name} doesn't hold a form");
    }
    catch (Exception ex) when (ex is OpenMcdf.FileFormatException or InvalidDataException or EndOfStreamException)
    {
      return (null, $"{name} can't be read: {ex.Message}");
    }

    var frame = string.Join("\r\n", lines.Where(l => !OleObjectBlob().IsMatch(l))) + "\r\n";
    return (new FormDesigner(frame, storage) { Frx = offset == 0 && size + FrxHeaderLength == frx.Length ? frx : null }, null);
  }

  /// <summary>Writes the designer into a project, as a storage named after the form, with its <c>VBFrame</c>.</summary>
  /// <param name="project">The storage holding the project's <c>VBA</c> storage and <c>PROJECT</c> stream.</param>
  /// <param name="name">The form's name.</param>
  /// <param name="encoding">The project's code page.</param>
  public void WriteInto(Storage project, string name, Encoding encoding)
  {
    ArgumentNullException.ThrowIfNull(project);
    ArgumentNullException.ThrowIfNull(encoding);
    if (project.ContainsEntry(name)) throw new InvalidDataException($"The project already has something called {name}, where the form's designer goes.");
    var designer = project.CreateStorage(name);
    using var source = RootStorage.Open(new MemoryStream(Storage));
    designer.CLSID = source.CLSID;
    CopyChildren(source, designer);
    Write(designer, FrameStream, encoding.GetBytes(Frame));
    CopyTimes(source, designer);
  }

  /// <summary>
  /// Whether two designers hold the same form: the same <c>VBFrame</c>, ignoring the <c>TypeInfoVer</c> a workbook adds,
  /// and the same streams, wherever they are in the compound file.
  /// </summary>
  public bool SameAs(FormDesigner other)
  {
    ArgumentNullException.ThrowIfNull(other);
    if (!WithoutTypeInfo(Frame).Equals(WithoutTypeInfo(other.Frame), StringComparison.Ordinal)) return false;
    var mine = Streams(Storage);
    var theirs = Streams(other.Storage);
    return mine.Count == theirs.Count && mine.All(s => theirs.TryGetValue(s.Key, out var bytes) && bytes.AsSpan().SequenceEqual(s.Value));
  }

  /// <summary>
  /// The <c>Begin</c> block as the VBA editor exports it: <c>VBFrame</c> without <c>TypeInfoVer</c>, and with the
  /// <c>OleObjectBlob</c> line naming the <c>.frx</c>, among the others in their order.
  /// </summary>
  public string ExportedFrame(string frxName)
  {
    var lines = Frame.Split("\r\n").Where(l => !TypeInfoVer().IsMatch(l)).ToList();
    var blob = $"   OleObjectBlob   =   \"{frxName}\":0000";
    var end = lines.FindLastIndex(l => l == "End");
    if (end < 0) end = lines.Count;
    var at = lines.FindIndex(1, l => Property().Match(l) is { Success: true } p && string.CompareOrdinal(p.Groups["name"].Value, "OleObjectBlob") > 0);
    lines.Insert(at < 1 || at > end ? end : at, blob);
    return string.Join("\r\n", lines).TrimEnd('\r', '\n') + "\r\n";
  }

  /// <summary>
  /// The <c>.frx</c>: the one it was read from, or one made of a header and the compound file. The header is <c>LB</c>,
  /// two bytes the VBA editor always writes as 8, the size of the rest, eight zeros, and what looks like the form's outer
  /// size, its client size with the borders and title bar the samples had (step S4a).
  /// </summary>
  public byte[] FrxBytes()
  {
    if (Frx is not null) return Frx;
    var header = new byte[FrxHeaderLength];
    header[0] = (byte)'L';
    header[1] = (byte)'B';
    header[2] = 8;
    BinaryPrimitives.WriteInt32LittleEndian(header.AsSpan(4), Storage.Length);
    BinaryPrimitives.WriteInt32LittleEndian(header.AsSpan(16), FrameNumber("ClientWidth") + 192);
    BinaryPrimitives.WriteInt32LittleEndian(header.AsSpan(20), FrameNumber("ClientHeight") + 512);
    return [.. header, .. Storage];
  }

  int FrameNumber(string property) =>
    Frame.Split("\r\n").Select(l => Property().Match(l)).FirstOrDefault(m => m.Success && m.Groups["name"].Value == property) is { } match
      && int.TryParse(match.Groups["value"].Value.Split('\'')[0].Trim(), out var value) ? value : 0;

  static string WithoutTypeInfo(string frame) => string.Join("\r\n", frame.Split("\r\n").Where(l => !TypeInfoVer().IsMatch(l))).TrimEnd();

  static Dictionary<string, byte[]> Streams(byte[] compoundFile)
  {
    var streams = new Dictionary<string, byte[]>(StringComparer.Ordinal);
    using var root = RootStorage.Open(new MemoryStream(compoundFile));
    void Walk(Storage storage, string path)
    {
      foreach (var entry in storage.EnumerateEntries())
      {
        if (entry.Type == EntryType.Storage) Walk(storage.OpenStorage(entry.Name), path + entry.Name + "/");
        else streams[path + entry.Name] = Read(storage, entry.Name);
      }
    }

    Walk(root, "");
    return streams;
  }

  static void CopyChildren(Storage from, Storage to, string? skip = null)
  {
    foreach (var entry in from.EnumerateEntries())
    {
      if (entry.Name == skip) continue;
      if (entry.Type == EntryType.Storage)
      {
        var source = from.OpenStorage(entry.Name);
        var child = to.CreateStorage(entry.Name);
        child.CLSID = source.CLSID;
        CopyChildren(source, child);
        CopyTimes(source, child);
      }
      else
      {
        Write(to, entry.Name, Read(from, entry.Name));
      }
    }
  }

  // A storage keeps when it was made and changed, which a copy would otherwise get as now, so the same form would give a
  // different .frx each time it's exported. A root has no creation time.
  static void CopyTimes(Storage from, Storage to)
  {
    if (to is not RootStorage) to.CreationTime = from.CreationTime;
    to.ModifiedTime = from.ModifiedTime;
  }

  static byte[] Read(Storage storage, string name)
  {
    using var stream = storage.OpenStream(name);
    var buffer = new MemoryStream();
    stream.CopyTo(buffer);
    return buffer.ToArray();
  }

  static void Write(Storage storage, string name, byte[] data)
  {
    using var stream = storage.CreateStream(name);
    stream.Write(data);
  }

  // From VERSION to the End that closes the form's Begin, before the Attribute lines.
  [GeneratedRegex(@"\AVERSION[^\r\n]*\r\nBegin\s.*?^End(?=\r\n|\z)", RegexOptions.Singleline | RegexOptions.Multiline | RegexOptions.IgnoreCase)]
  private static partial Regex BeginBlock();

  [GeneratedRegex(@"^\s*OleObjectBlob\s*=\s*""(?<file>[^""]+)"":(?<offset>[0-9A-Fa-f]+)", RegexOptions.IgnoreCase)]
  private static partial Regex OleObjectBlob();

  [GeneratedRegex(@"^\s*TypeInfoVer\s*=", RegexOptions.IgnoreCase)]
  private static partial Regex TypeInfoVer();

  [GeneratedRegex(@"^\s+(?<name>\w+)\s*=\s*(?<value>.*)$")]
  private static partial Regex Property();
}
