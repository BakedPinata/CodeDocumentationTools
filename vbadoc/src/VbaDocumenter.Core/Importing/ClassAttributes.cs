using System.Text.RegularExpressions;
using VbaDocumenter.Core.Extraction;

namespace VbaDocumenter.Core.Importing;

/// <summary>What a class's attribute lines have to be changed to, and how many characters of the source they replace.</summary>
/// <param name="Length">The length of the attribute lines at the top of the source, which <paramref name="Header"/> replaces.</param>
/// <param name="Header">The attribute lines as a VBA project stores them, each ending in a line break.</param>
public sealed record ClassHeader(int Length, string Header);

/// <summary>
/// A class's or a UserForm's attributes as a VBA project stores them. The VBA editor's File > Export leaves three out
/// (<c>VB_Base</c>, <c>VB_TemplateDerived</c> and <c>VB_Customizable</c>), and its File > Import puts them back. A
/// project written straight from an exported file has none of them, and Excel can't load it: for a class, "Invalid
/// procedure call or argument", with no line shown, and for a form, a crash (step S4a). So a class or form going into a
/// project gets what's missing, as DECISIONS.md says.
/// </summary>
public static partial class ClassAttributes
{
  /// <summary>The base class id every class module has, in <c>VB_Base</c>.</summary>
  public const string ClassBase = "\"0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}\"";

  // The attributes a project stores at the top of a class, in the order Excel writes them, with what a new class has.
  // VB_Name has no default: a source without one is left as it is.
  static readonly (string Name, Func<string>? Value)[] ClassStored =
  [
    ("VB_Name", null),
    ("VB_Base", () => ClassBase),
    ("VB_GlobalNameSpace", () => "False"),
    ("VB_Creatable", () => "False"),
    ("VB_PredeclaredId", () => "False"),
    ("VB_Exposed", () => "False"),
    ("VB_TemplateDerived", () => "False"),
    ("VB_Customizable", () => "False"),
  ];

  // A form's are the same but for VB_Base, which names two GUIDs of the form's own (FormBase), and VB_PredeclaredId.
  static readonly (string Name, Func<string>? Value)[] FormStored =
  [
    ("VB_Name", null),
    ("VB_Base", FormBase),
    ("VB_GlobalNameSpace", () => "False"),
    ("VB_Creatable", () => "False"),
    ("VB_PredeclaredId", () => "True"),
    ("VB_Exposed", () => "False"),
    ("VB_TemplateDerived", () => "False"),
    ("VB_Customizable", () => "False"),
  ];

  /// <summary>
  /// A new <c>VB_Base</c> for a form: two GUIDs, which the VBA editor makes on import. Excel makes new ones when it
  /// compiles the project, so any will do (step S4a).
  /// </summary>
  public static string FormBase() => $"\"0{{{NewGuid()}}}{{{NewGuid()}}}\"";

  static string NewGuid() => Guid.NewGuid().ToString().ToUpperInvariant();

  /// <summary>
  /// A module's source as a project stores it: a class's or form's with any attributes it's missing, and anything else
  /// as it is.
  /// </summary>
  public static string StoredSource(VbaModule module)
  {
    ArgumentNullException.ThrowIfNull(module);
    return module.Kind is VbaModuleKind.Class or VbaModuleKind.Form ? Complete(module.Source, module.Kind) : module.Source;
  }

  /// <summary>
  /// A module's source as it's compared with another's: as a project stores it, and for a form without its
  /// <c>VB_Base</c>, which differs from one copy of a form to the next.
  /// </summary>
  public static string Comparable(VbaModule module)
  {
    ArgumentNullException.ThrowIfNull(module);
    var stored = StoredSource(module);
    return module.Kind == VbaModuleKind.Form ? BaseLine().Replace(stored, "") : stored;
  }

  /// <summary>A class's source with any attributes it's missing, or the same string when it has them all.</summary>
  public static string Complete(string source) => Complete(source, VbaModuleKind.Class);

  /// <summary>A class's or form's source with any attributes it's missing, or the same string when it has them all.</summary>
  public static string Complete(string source, VbaModuleKind kind)
  {
    ArgumentNullException.ThrowIfNull(source);
    return Completion(source, kind) is { } change ? change.Header + source[change.Length..] : source;
  }

  /// <summary>How a class's attribute lines have to change: see <see cref="Completion(string, VbaModuleKind)"/>.</summary>
  public static ClassHeader? Completion(string source) => Completion(source, VbaModuleKind.Class);

  /// <summary>
  /// How a class's or form's attribute lines have to change, or null when nothing's missing, the source doesn't start
  /// with its <c>Attribute VB_Name</c> line, or it's another kind of module. The ones it has keep their values, in
  /// Excel's order, with any others after them, so only the lines at the top change.
  /// </summary>
  public static ClassHeader? Completion(string source, VbaModuleKind kind)
  {
    ArgumentNullException.ThrowIfNull(source);
    var stored = kind switch
    {
      VbaModuleKind.Class => ClassStored,
      VbaModuleKind.Form => FormStored,
      _ => null,
    };
    if (stored is null) return null;

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
    if (!has.Contains("VB_Name") || stored.All(s => has.Contains(s.Name))) return null;

    lineBreak ??= "\r\n";
    var header = new System.Text.StringBuilder();
    foreach (var (name, value) in stored)
    {
      var text = lines.FirstOrDefault(l => l.Name.Equals(name, StringComparison.OrdinalIgnoreCase)).Text ?? $"Attribute {name} = {value!()}";
      header.Append(text).Append(lineBreak);
    }

    foreach (var (name, text) in lines.Where(l => !stored.Any(s => s.Name.Equals(l.Name, StringComparison.OrdinalIgnoreCase))))
      header.Append(text).Append(lineBreak);

    return new ClassHeader(at, header.ToString());
  }

  // One module attribute line at the given position, with its line break if it has one.
  [GeneratedRegex(@"\G(?<text>Attribute[ \t]+(?<name>VB_\w+)[ \t]*=[^\r\n]*)(?<break>\r\n|\n|\r|$)", RegexOptions.IgnoreCase)]
  private static partial Regex AttributeLine();

  [GeneratedRegex(@"^Attribute[ \t]+VB_Base[ \t]*=[^\r\n]*(\r\n|\n|\r)?", RegexOptions.IgnoreCase | RegexOptions.Multiline)]
  private static partial Regex BaseLine();
}
