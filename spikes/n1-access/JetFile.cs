using System.Buffers.Binary;
using System.Text;

namespace AccessSpike;

/// <summary>
/// Choice C: just enough of the Access database file format (ACE, the .accdb successor of Jet 4) to read a table's
/// rows with no Access and no database engine. The format isn't published; this follows what the mdbtools project
/// worked out. It reads 4 KB pages, a table's definition, its rows wherever they are in the file, and long values
/// (memo and OLE fields) that sit in the row, on a page of their own, or chained over several.
/// </summary>
public sealed class JetFile
{
  private const int PageSize = 4096;
  private readonly byte[] bytes;

  public JetFile(byte[] bytes)
  {
    // Page 0 says which format: 0 is Jet 3 (.mdb, Access 97), 1 Jet 4, 2 and up ACE (.accdb).
    if (bytes.Length < PageSize * 3 || bytes[0] != 0 || bytes[1] != 1 || bytes[2] != 0 || bytes[3] != 0)
    {
      throw new InvalidDataException("This isn't an Access database file.");
    }

    Version = bytes[0x14];
    if (Version < 1)
    {
      throw new InvalidDataException("This is an Access 97 database, which this doesn't read.");
    }

    this.bytes = bytes;
  }

  /// <summary>The file's format: 1 is Jet 4, 2 ACE 12 (Access 2007), 3 ACE 14 and so on.</summary>
  public int Version { get; }

  public int PageCount => bytes.Length / PageSize;

  private ReadOnlySpan<byte> Page(int number) => bytes.AsSpan(number * PageSize, PageSize);

  /// <summary>A column of a table: where its value is in a row.</summary>
  public sealed record Column(string Name, int Type, int Number, int VarIndex, int FixedOffset, int Length, bool IsFixed);

  /// <summary>A table's definition.</summary>
  public sealed record Table(int DefinitionPage, IReadOnlyList<Column> Columns);

  /// <summary>The tables in the file, by name, each with the page its definition starts on.</summary>
  public Dictionary<string, int> Tables()
  {
    // MSysObjects, the catalogue, is always defined on page 2.
    var catalogue = ReadTable(2);
    var tables = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
    foreach (var row in Rows(catalogue))
    {
      // Type 1 is a local table. The low three bytes of its Id are its definition's page.
      if (row["Type"] is short type && type == 1 && row["Name"] is string name && row["Id"] is int id)
      {
        tables[name] = id & 0x00FFFFFF;
      }
    }

    return tables;
  }

  public Table ReadTable(int definitionPage)
  {
    // A definition can run over several pages, each with an 8-byte header, the first four bytes of the next page's
    // number after the page type.
    var definition = new List<byte>();
    for (var page = definitionPage; page != 0;)
    {
      var span = Page(page);
      if (span[0] != 0x02)
      {
        throw new InvalidDataException($"Page {page} isn't a table definition.");
      }

      definition.AddRange(definition.Count == 0 ? span : span[8..]);
      page = BinaryPrimitives.ReadInt32LittleEndian(span[4..]);
    }

    var d = definition.ToArray().AsSpan();
    int columnCount = BinaryPrimitives.ReadInt16LittleEndian(d[45..]);
    var realIndexes = BinaryPrimitives.ReadInt32LittleEndian(d[51..]);
    var at = 63 + realIndexes * 12;
    var parts = new List<(int Type, int Number, int VarIndex, int FixedOffset, int Length, bool IsFixed)>();
    for (var i = 0; i < columnCount; i++, at += 25)
    {
      parts.Add((d[at], BinaryPrimitives.ReadInt16LittleEndian(d[(at + 5)..]), BinaryPrimitives.ReadInt16LittleEndian(d[(at + 7)..]),
        BinaryPrimitives.ReadInt16LittleEndian(d[(at + 21)..]), BinaryPrimitives.ReadInt16LittleEndian(d[(at + 23)..]), (d[at + 15] & 0x01) != 0));
    }

    var columns = new List<Column>();
    foreach (var (type, number, varIndex, fixedOffset, length, isFixed) in parts)
    {
      int nameLength = BinaryPrimitives.ReadInt16LittleEndian(d[at..]);
      var name = Encoding.Unicode.GetString(d.Slice(at + 2, nameLength));
      at += 2 + nameLength;
      columns.Add(new Column(name, type, number, varIndex, fixedOffset, length, isFixed));
    }

    return new Table(definitionPage, columns);
  }

  /// <summary>
  /// Every row of a table, as column name to value. The file is searched for the table's data pages rather than
  /// following its map of them, which is enough for a spike and reads the whole file once.
  /// </summary>
  public IEnumerable<Dictionary<string, object?>> Rows(Table table)
  {
    for (var number = 1; number < PageCount; number++)
    {
      if (Page(number)[0] != 0x01 || BinaryPrimitives.ReadInt32LittleEndian(Page(number)[4..]) != table.DefinitionPage)
      {
        continue;
      }

      int rowCount = BinaryPrimitives.ReadInt16LittleEndian(Page(number)[12..]);
      for (var i = 0; i < rowCount; i++)
      {
        var (start, end, flags) = RowBounds(number, i);
        // 0x4000 is a deleted row, and 0x8000 one that only points to where the row moved to, which is read there.
        if ((flags & 0xC000) != 0 || end - start < 4)
        {
          continue;
        }

        yield return ReadRow(table, bytes.AsSpan(number * PageSize + start, end - start));
      }
    }
  }

  private (int Start, int End, int Flags) RowBounds(int page, int row)
  {
    var span = Page(page);
    int offset = BinaryPrimitives.ReadUInt16LittleEndian(span[(14 + row * 2)..]);
    var end = row == 0 ? PageSize : BinaryPrimitives.ReadUInt16LittleEndian(span[(14 + (row - 1) * 2)..]) & 0x1FFF;
    return (offset & 0x1FFF, end, offset & 0xE000);
  }

  private Dictionary<string, object?> ReadRow(Table table, ReadOnlySpan<byte> row)
  {
    // A row: how many columns it has, the fixed-length values, the variable-length values, a table of where those
    // start (written backwards from the end), how many of them there are, and a bit for each column that isn't null.
    int rowColumns = BinaryPrimitives.ReadInt16LittleEndian(row);
    var maskSize = (rowColumns + 7) / 8;
    var mask = row[^maskSize..];
    int varCount = BinaryPrimitives.ReadInt16LittleEndian(row[(row.Length - maskSize - 2)..]);
    var varOffsets = new int[Math.Max(0, varCount) + 1];
    for (var i = 0; i < varOffsets.Length; i++)
    {
      varOffsets[i] = BinaryPrimitives.ReadInt16LittleEndian(row[(row.Length - maskSize - 4 - i * 2)..]);
    }

    var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
    foreach (var column in table.Columns)
    {
      var present = column.Number < rowColumns && (mask[column.Number / 8] & (1 << (column.Number % 8))) != 0;
      if (column.Type == 1)
      {
        values[column.Name] = present; // a Yes/No is only its bit
        continue;
      }

      if (!present)
      {
        values[column.Name] = null;
        continue;
      }

      ReadOnlySpan<byte> data;
      if (column.IsFixed)
      {
        data = row.Slice(2 + column.FixedOffset, column.Length);
      }
      else if (column.VarIndex < varCount)
      {
        data = row[varOffsets[column.VarIndex]..varOffsets[column.VarIndex + 1]];
      }
      else
      {
        values[column.Name] = null;
        continue;
      }

      values[column.Name] = column.Type switch
      {
        2 => data[0],
        3 => BinaryPrimitives.ReadInt16LittleEndian(data),
        4 => BinaryPrimitives.ReadInt32LittleEndian(data),
        10 => Text(data),
        11 => LongValue(data),
        12 => Text(LongValue(data)),
        _ => data.ToArray(),
      };
    }

    return values;
  }

  /// <summary>
  /// Text is UTF-16, or, after the mark FF FE, "compressed": a byte a character, switching to two bytes a character
  /// and back at each zero byte.
  /// </summary>
  private static string Text(ReadOnlySpan<byte> data)
  {
    if (data.Length < 2 || data[0] != 0xFF || data[1] != 0xFE)
    {
      return Encoding.Unicode.GetString(data);
    }

    var text = new StringBuilder();
    var oneByte = true;
    for (var i = 2; i < data.Length;)
    {
      if (data[i] == 0)
      {
        oneByte = !oneByte;
        i++;
      }
      else if (oneByte)
      {
        text.Append((char)data[i++]);
      }
      else if (i + 1 < data.Length)
      {
        text.Append((char)(data[i] | (data[i + 1] << 8)));
        i += 2;
      }
      else
      {
        break;
      }
    }

    return text.ToString();
  }

  /// <summary>
  /// A memo or OLE value. The field holds its length and where it is: in the field itself, in one row of another
  /// page, or in a chain of rows, each starting with a pointer to the next.
  /// </summary>
  private byte[] LongValue(ReadOnlySpan<byte> field)
  {
    var lengthAndFlags = BinaryPrimitives.ReadUInt32LittleEndian(field);
    var length = (int)(lengthAndFlags & 0x3FFFFFFF);
    if ((lengthAndFlags & 0x80000000) != 0)
    {
      return field.Slice(12, length).ToArray();
    }

    var pointer = BinaryPrimitives.ReadInt32LittleEndian(field[4..]);
    if ((lengthAndFlags & 0x40000000) != 0)
    {
      var (start, end, _) = RowBounds(pointer >> 8, pointer & 0xFF);
      return Page(pointer >> 8)[start..end].ToArray();
    }

    var value = new List<byte>(length);
    while (pointer != 0 && value.Count < length)
    {
      var (start, end, _) = RowBounds(pointer >> 8, pointer & 0xFF);
      var chunk = Page(pointer >> 8)[start..end];
      pointer = BinaryPrimitives.ReadInt32LittleEndian(chunk);
      value.AddRange(chunk[4..]);
    }

    return [.. value];
  }
}
