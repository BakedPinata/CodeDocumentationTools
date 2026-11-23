using System.Buffers.Binary;
using OpenMcdf;

namespace S4Spike;

/// <summary>
/// Replaces the JPEG an Image control holds in a .frx with another, padded to exactly the same length with comment
/// segments, so every stream, length and offset in the .frx stays as the VBA editor wrote it. Used to take a picture that
/// isn't ours out of the sample before it became a fixture (step S4b).
/// </summary>
static class ImageReplacer
{
  public static int Run(string frxPath, string jpegPath, string outPath)
  {
    var frx = File.ReadAllBytes(frxPath);
    var size = BinaryPrimitives.ReadInt32LittleEndian(frx.AsSpan(4));
    var buffer = new MemoryStream();
    buffer.Write(frx, 24, size);
    buffer.Position = 0;
    using (var root = RootStorage.Open(buffer, StorageModeFlags.LeaveOpen))
    {
      var jpeg = File.ReadAllBytes(jpegPath);
      int replaced = 0;
      foreach (var (storage, name, path) in AllStreams(root, ""))
      {
        var data = Streams.Read(storage, name);
        bool changed = false;
        for (int at = data.AsSpan().IndexOf(JpegStart); at > 4; at = Next(data, at))
        {
          // A picture property is its length, then its bytes; a JPEG starts FF D8 FF.
          int length = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(at - 4));
          if (length <= 0 || at + length > data.Length) continue;
          Padded(jpeg, length).CopyTo(data, at);
          Console.WriteLine($"Replaced a JPEG of {length} bytes in {path}");
          changed = true;
          replaced++;
        }

        if (!changed) continue;
        using var stream = storage.OpenStream(name);
        stream.Write(data);
      }

      if (replaced == 0) throw new InvalidDataException("No stream holds a JPEG.");
    }

    var compound = buffer.ToArray();
    if (compound.Length != size) throw new InvalidDataException($"The compound file changed size, {size} to {compound.Length}.");
    File.WriteAllBytes(outPath, [.. frx.AsSpan(0, 24), .. compound]);
    return 0;
  }

  static readonly byte[] JpegStart = [0xFF, 0xD8, 0xFF];

  static int Next(byte[] data, int after)
  {
    int found = data.AsSpan(after + 1).IndexOf(JpegStart);
    return found < 0 ? -1 : after + 1 + found;
  }

  static IEnumerable<(Storage Storage, string Name, string Path)> AllStreams(Storage storage, string path)
  {
    foreach (var entry in storage.EnumerateEntries().ToList())
    {
      if (entry.Type == EntryType.Storage)
      {
        foreach (var stream in AllStreams(storage.OpenStorage(entry.Name), path + Inspector.Printable(entry.Name) + "/")) yield return stream;
      }
      else
      {
        yield return (storage, entry.Name, path + Inspector.Printable(entry.Name));
      }
    }
  }

  // Comment segments (FF FE, a 2-byte length that counts itself, then any bytes) straight after the start of the image.
  static byte[] Padded(byte[] jpeg, int length)
  {
    var output = new List<byte>(length) { 0xFF, 0xD8 };
    int needed = length - jpeg.Length;
    if (needed < 0) throw new InvalidDataException($"The new JPEG is {jpeg.Length} bytes, longer than the {length} it replaces.");
    while (needed > 0)
    {
      int segment = Math.Min(needed, 65535 + 2);
      if (needed - segment is > 0 and < 4) segment -= 4; // leave enough for one more whole segment
      output.AddRange([0xFF, 0xFE, (byte)((segment - 2) >> 8), (byte)((segment - 2) & 0xFF)]);
      output.AddRange(new byte[segment - 4]);
      needed -= segment;
    }

    output.AddRange(jpeg.AsSpan(2).ToArray());
    if (output.Count != length) throw new InvalidDataException($"Padding came to {output.Count}, not {length}.");
    return [.. output];
  }
}
