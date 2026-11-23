using System.Text.RegularExpressions;
using VbaDocumenter.Core.Extraction;

namespace VbaDocumenter.Core.Importing;

/// <summary>What a class's attribute lines have to be changed to, and how many characters of the source they replace.</summary>
/// <param name="Length">The length of the attribute lines at the top of the source, which <paramref name="Header"/> replaces.</param>
/// <param name="Header">The attribute lines as a VBA project stores them, each ending in a line break.</param>
public sealed record ClassHeader(int Length, string Header);

/// <summary>
/// A class's attributes as a VBA project stores them. The VBA editor's File > Export leaves three out (<c>VB_Base</c>,
/// <c>VB_TemplateDerived</c> and <c>VB_Customizable</c>), and its File > Import puts them back. A project written
/// straight from an exported file has none of them, and Excel can't load it: "Invalid procedure call or argument", with
/// no line shown. So a class going into a project gets what's missing, as DECISIONS.md says.
/// </summary>
public static partial class ClassAttributes
{
  /// <summary>The base class id every class module has, in <c>VB_Base</c>.</summary>
  public const string ClassBase = "\"0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}\"";

  // The attributes a project stores at the top of a class, in the order Excel writes them, with what a new class has.
  // VB_Name has no default: a source without one is left as it is.
  static readonly (string Name, string? Value)[] Stored =
  [
    ("VB_Name", null),
    ("VB_Base", ClassBase),
    ("VB_GlobalNameSpace", "False"),
    ("VB_Creatable", "False"),
    ("VB_PredeclaredId", "False"),
    ("VB_Exposed", "False"),
    ("VB_TemplateDerived", "False"),
    ("VB_Customizable", "False"),
  ];

  /// <summary>
  /// A module's source as a project stores it: a class's with any attributes it's missing, and anything else as it is.
  /// </summary>
  public static string StoredSource(VbaModule module)
  {
    ArgumentNullException.ThrowIfNull(module);
    return module.Kind == VbaModuleKind.Class ? Complete(module.Source) : module.Source;
  }

  /// <summary>A class's source with any attributes it's missing, or the same string when it has them all.</summary>
  public static string Complete(string source)
  {
    ArgumentNullException.ThrowIfNull(source);
    return Completion(source) is { } change ? change.Header + source[change.Length..] : source;
  }

  /// <summary>
  /// How a class's attribute lines have to change, or null when nothing's missing or the source doesn't start with its
  /// <c>Attribute VB_Name</c> line. The ones it has keep their values, in Excel's order, with any others after them, so
  /// only the lines at the top change.
  /// </summary>
  public static ClassHeader? Completion(string source)
  {
    ArgumentNullException.ThrowIfNull(source);
    var lines = new List<(string Name, string Text)>();
    var at = 0;
    string? lineBreak = null;
    while (AttributeLine().Match(source, at) is { Success: true } line)
    {
      lines.Add((line.Groups["name"].Value, line.Groups["text"].Value));
      lineBreak ??= line.Groups["break"].Length > 0 ? line.Groups["break"].Value : null;
      at += line.Length;
      if (line.Groups["break"].Length == 0) break;
    }

    var has = lines.Select(l => l.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
    if (!has.Contains("VB_Name") || Stored.All(s => has.Contains(s.Name))) return null;

    lineBreak ??= "\r\n";
    var header = new System.Text.StringBuilder();
    foreach (var (name, value) in Stored)
    {
      var text = lines.FirstOrDefault(l => l.Name.Equals(name, StringComparison.OrdinalIgnoreCase)).Text ?? $"Attribute {name} = {value}";
      header.Append(text).Append(lineBreak);
    }

    foreach (var (name, text) in lines.Where(l => !Stored.Any(s => s.Name.Equals(l.Name, StringComparison.OrdinalIgnoreCase))))
      header.Append(text).Append(lineBreak);

    return new ClassHeader(at, header.ToString());
  }

  // One module attribute line at the given position, with its line break if it has one.
  [GeneratedRegex(@"\G(?<text>Attribute[ \t]+(?<name>VB_\w+)[ \t]*=[^\r\n]*)(?<break>\r\n|\n|\r|$)", RegexOptions.IgnoreCase)]
  private static partial Regex AttributeLine();
}
