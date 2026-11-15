using CodeDocumenter.CodeViews;
using VbaDocumenter.Core.Analysis;
using VbaDocumenter.Core.Extraction;
using VbaDocumenter.Core.Parsing;

namespace VbaDocumenter.Core.Output;

/// <summary>
/// Works out the class diagram's box contents and relationships from the parsed project, as the code views library
/// draws them (<see cref="CodeClassLine"/>, <see cref="CodeClassEdge"/>): how VBA's members read as lines of a box.
/// </summary>
public static class ClassDiagram
{
  const int MaxValueLength = 40;

  /// <summary>
  /// The line shown for each member of a module. Members without a line aren't shown: enum values (the Enum is),
  /// the second and third accessors of a property, which is shown once on its first accessor, and a declaration
  /// repeated that would give exactly the same line. (A name declared again under #If is already one member, with its
  /// types joined: "-m_Dictionary : Scripting.Dictionary | Object".)
  /// </summary>
  public static Dictionary<Member, CodeClassLine> Lines(ModuleModel module)
  {
    var lines = new Dictionary<Member, CodeClassLine>();
    var propertiesSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    var linesSeen = new HashSet<CodeClassLine>();
    foreach (var m in module.Members)
    {
      if (m.Kind == MemberKind.EnumMember) continue;
      CodeClassLine line;
      if (IsProperty(m))
      {
        if (!propertiesSeen.Add(m.Name)) continue;
        line = PropertyLine(m, module.Members.Where(o => IsProperty(o) && string.Equals(o.Name, m.Name, StringComparison.OrdinalIgnoreCase)).ToList());
      }
      else
      {
        line = Line(m);
      }
      if (linesSeen.Add(line)) lines[m] = line;
    }
    return lines;
  }

  static bool IsProperty(Member m) => m.Kind is MemberKind.PropertyGet or MemberKind.PropertyLet or MemberKind.PropertySet;

  static string Visibility(Member m) => m.Visibility switch
  {
    Parsing.Visibility.Private => "-",
    Parsing.Visibility.Friend => "~",
    _ => "+",
  };

  static string? TypeText(string? type, bool isArray) => type == null ? null : type + (isArray ? "()" : "");

  static CodeClassLine Line(Member m)
  {
    string v = Visibility(m);
    string parameters = MemberFormatter.Parameters(m);
    string? returnType = TypeText(MemberFormatter.TypeText(m), m.IsArray);
    return m.Kind switch
    {
      MemberKind.Sub => new($"{v}{m.Name}", parameters),
      MemberKind.Function => new($"{v}{m.Name}", parameters, returnType),
      MemberKind.Event => new($"+Event {m.Name}", parameters),
      MemberKind.Declare => new($"{v}Declare {(m.IsFunctionDeclare ? "Function" : "Sub")} {m.Name}", parameters, returnType),
      MemberKind.Variable => new($"{v}{(m.IsWithEvents ? "WithEvents " : "")}{m.Name}{(m.IsArray ? "()" : "")}" +
                                 (MemberFormatter.TypeText(m) is { } variableType ? $" : {(m.IsAutoNew ? "New " : "")}{variableType}" : "")),
      MemberKind.Constant => new($"{v}Const {m.Name}{(MemberFormatter.TypeText(m) is { } constantType ? $" : {constantType}" : "")}" +
                                 (m.Value == null ? "" : $" = {Shorten(m.Value)}")),
      MemberKind.Enum => new($"{v}Enum {m.Name}"),
      MemberKind.Type => new($"{v}Type {m.Name}"),
      _ => new($"{v}{m.Name}"),
    };
  }

  // A property is shown once, with the accessors it has: "+Property Status : String «Get/Let»". The type is
  // the Get's return type, or else the value parameter of the Let or Set. An indexed property (one with
  // parameters other than the value) is drawn as a method so its parameters can be shown.
  static CodeClassLine PropertyLine(Member first, List<Member> accessors)
  {
    var get = accessors.FirstOrDefault(a => a.Kind == MemberKind.PropertyGet);
    var setter = accessors.FirstOrDefault(a => a.Kind != MemberKind.PropertyGet);
    string? type = get != null ? TypeText(MemberFormatter.TypeText(get), get.IsArray) : setter?.Parameters.LastOrDefault()?.Type;
    string kinds = string.Join("/", new[] { MemberKind.PropertyGet, MemberKind.PropertyLet, MemberKind.PropertySet }
        .Where(k => accessors.Any(a => a.Kind == k))
        .Select(k => k.ToString()["Property".Length..]));
    string text = $"{Visibility(first)}Property {first.Name}";

    // The index parameters: all of a Get's, or all but the last (the value) of a Let's or Set's.
    var index = get != null ? get.Parameters : setter!.Parameters.Take(Math.Max(0, setter.Parameters.Count - 1)).ToList();
    if (index.Count > 0)
      return new(text, MemberFormatter.Parameters(index), $"{type ?? "Variant"} «{kinds}»");
    return new($"{text}{(type == null ? "" : $" : {type}")} «{kinds}»");
  }

  static string Shorten(string value) => value.Length <= MaxValueLength ? value : value[..(MaxValueLength - 1)] + "…";

  /// <summary>
  /// The diagram's relationships, in the module edges' order. Implements edges become realisations. A module-level
  /// variable typed as another module becomes an association labelled with the variable's name (marked for
  /// WithEvents). Every other module edge becomes a dependency, unless an association already links the two modules.
  /// </summary>
  public static List<CodeClassEdge> Edges(DependencyModel model)
  {
    var modulesByName = model.Modules.ToDictionary(m => m.Id, StringComparer.OrdinalIgnoreCase);
    var associations = new List<CodeClassEdge>();
    foreach (var node in model.Modules)
    {
      foreach (var v in node.Module.Members.Where(m => m.Kind == MemberKind.Variable))
      {
        // A variable declared again under #If holds whichever of its types the project class is, if any.
        foreach (var type in v.Types)
        {
          if (type == null) continue;
          var typeName = type[(type.LastIndexOf('.') + 1)..]; // "VBAProject.Order" names the same module as "Order"
          if (!modulesByName.TryGetValue(typeName, out var target) || target.Id == node.Id) continue;
          if (target.Module.Kind == VbaModuleKind.Standard) continue; // not a type you can hold
          associations.Add(new CodeClassEdge(node.Id, target.Id, "association", (v.IsWithEvents ? "WithEvents " : "") + v.Name));
          break;
        }
      }
    }

    var associated = associations.Select(a => (a.From, a.To)).ToHashSet();
    var edges = new List<CodeClassEdge>();
    var dependencies = new HashSet<(string, string)>();
    foreach (var e in model.ModuleEdges)
    {
      if (e.Kind == EdgeKind.Implements)
        edges.Add(new CodeClassEdge(e.From, e.To, "realization", null));
      else if (!associated.Contains((e.From, e.To)) && dependencies.Add((e.From, e.To)))
        edges.Add(new CodeClassEdge(e.From, e.To, "dependency", null));
    }
    edges.AddRange(associations);
    return edges;
  }
}
