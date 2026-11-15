using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace CodeDocumenter.CodeViews;

// The code model: the page's data, in terms no language owns. A language's Core fills it in, and the views draw it.
// Files are apart from what's in them: a unit or member says which file and lines it's in, so one file can hold
// several units and one unit can be spread over several files. Ids are names, and favourites, hidden items, the
// filter and saved views are kept by id. Property names become camelCase in the JSON, and nulls are left out.

/// <summary>What the page documents.</summary>
/// <param name="Title">Its name, as the page shows it: a workbook's file name, a solution's.</param>
/// <param name="SourceMode"><c>embedded</c> (the page holds the source), <c>folder</c> (Load folder reads it) or <c>none</c>.</param>
public sealed record CodeSource(string Title, string SourceMode);

/// <summary>A file of source.</summary>
/// <param name="Id">What units and members call it. Its name will do when names are unique.</param>
/// <param name="Name">Its file name, which Load folder looks for.</param>
/// <param name="Hash">The first 16 hex digits of its SHA-256, so the page can tell whether a loaded folder still matches.</param>
/// <param name="Lines">How many lines it has.</param>
public sealed record CodeFile(string Id, string Name, string Hash, int Lines);

/// <summary>A group of units, shown as a folder: a namespace, a package, a <c>'@Folder</c>. Groups nest to any depth.</summary>
/// <param name="Id">What units and other groups call it.</param>
/// <param name="Name">Its own name, without its parents'.</param>
/// <param name="Parent">The group it's in, or null at the top.</param>
public sealed record CodeGroup(string Id, string Name, string? Parent = null);

/// <summary>Where part of a unit is, for a unit that isn't the whole of one file.</summary>
public sealed record CodePart(string File, int First, int Last);

/// <summary>A unit: a module, a class, a type. The graph's upper level draws it as a node and the class diagram as a box.</summary>
/// <param name="Id">Its name, unique among units.</param>
/// <param name="Kind">One of the language's <see cref="CodeUnitKind"/>s.</param>
/// <param name="File">The file it's in, or the first of them.</param>
/// <param name="Group">The <see cref="CodeGroup"/> it's in, or null.</param>
/// <param name="Parts">
/// For a unit that isn't the whole of one file: where each part of it is, in order. Null when the unit is its file.
/// </param>
/// <param name="KindLabel">What to call its kind when the kind's own label doesn't say it: <c>Predeclared class</c>.</param>
/// <param name="Note">A word about it for the explorer, beside its name: <c>Predeclared</c>.</param>
/// <param name="UsedBy">How many other units use it, or null for none.</param>
/// <param name="Description">What it says about itself, or null.</param>
/// <param name="Extra">Anything more the tool's own script wants of it. The views don't read it.</param>
public sealed record CodeUnit(
  string Id, string Kind, string File, string? Group = null, IReadOnlyList<CodePart>? Parts = null, string? KindLabel = null,
  string? Note = null, int? UsedBy = null, string? Description = null, object? Extra = null);

/// <summary>
/// One line in a class diagram box, as plain text, which the page draws as it is. A line with
/// <see cref="Parameters"/> (possibly empty) is drawn as a method, e.g. "+Describe(x As Long) : String"; one without
/// is drawn as an attribute, e.g. "-m_Id : String". Attributes and methods go in separate compartments.
/// </summary>
/// <param name="Text">Visibility and name, plus the type for attributes.</param>
/// <param name="Parameters">For methods, the parameter list without the brackets.</param>
/// <param name="Type">For methods, the return type.</param>
public sealed record CodeClassLine(string Text, string? Parameters = null, string? Type = null);

/// <summary>A member of a unit.</summary>
/// <param name="Id">Its name, unique among members and units: the unit's id, a dot and its own name will do.</param>
/// <param name="Unit">The unit it's in.</param>
/// <param name="Name">Its own name.</param>
/// <param name="Kind">One of the language's <see cref="CodeMemberKind"/>s.</param>
/// <param name="Visibility"><c>public</c>, <c>private</c> or the language's own word for another.</param>
/// <param name="Signature">Its declaration on one line, shown as code.</param>
/// <param name="File">The file it's in, which needn't be its unit's first.</param>
/// <param name="Start">Its first line in that file, with the comments above it.</param>
/// <param name="End">Its last line.</param>
/// <param name="Declaration">The line it's declared on.</param>
/// <param name="Role">One of the language's <see cref="CodeRole"/>s, or <c>none</c>.</param>
/// <param name="RoleNote">Something to add to the role's badge for this member, or null.</param>
/// <param name="Unreferenced">Whether nothing in the code refers to it.</param>
/// <param name="Parent">The name of the member of the same unit it belongs under, as an Enum's value does, or null.</param>
/// <param name="Class">Its line in the class diagram, or null if it isn't shown there.</param>
/// <param name="UsedBy">How many other units use it, or null for none.</param>
/// <param name="Description">What it says about itself, or null.</param>
public sealed record CodeMember(
  string Id, string Unit, string Name, string Kind, string Visibility, string Signature, string File,
  int Start, int End, int Declaration, string Role = "none", string? RoleNote = null, bool Unreferenced = false,
  string? Parent = null, CodeClassLine? Class = null, int? UsedBy = null, string? Description = null);

/// <summary>An edge between two units, or two members (a unit's id stands for references to the unit itself).</summary>
/// <param name="Kind">One of the language's <see cref="CodeRelationship"/>s.</param>
/// <param name="Count">How many references it stands for.</param>
/// <param name="Lines">For a member edge, the lines of <paramref name="From"/>'s file where the references are; null for unit edges.</param>
public sealed record CodeEdge(string From, string To, string Kind, int Count, IReadOnlyList<int>? Lines = null);

/// <summary>A relationship between two boxes in the class diagram.</summary>
/// <param name="Kind">One of the language's <see cref="CodeClassRelationship"/>s.</param>
/// <param name="Label">What the line is labelled with, such as the variable that holds the other, or null.</param>
public sealed record CodeClassEdge(string From, string To, string Kind, string? Label = null);

/// <summary>A call that can't be followed: what's written, and where.</summary>
public sealed record CodeUnfollowed(string From, string Expression, int Line);

/// <summary>
/// A call, for the Calls tab, in the order of the code: what's called, or the expression when it can't be followed,
/// the line, the innermost block it's in, or null, and for a call made by a block's own statement, that block.
/// </summary>
public sealed record CodeCall(string From, string? To, string? Unfollowed, int Line, int? Block = null, int? Condition = null);

/// <summary>A block statement of a member (an If, a loop), for the Calls tab.</summary>
/// <param name="Id">Its number within <paramref name="From"/>, counted from 0.</param>
/// <param name="Kind">
/// What its own statement works out, which a call made there is marked with: <c>if</c>, <c>elseIf</c>, <c>do</c>,
/// <c>while</c>, <c>select</c>, <c>case</c>, <c>for</c>, <c>with</c>, or anything else for a plain condition.
/// </param>
/// <param name="Text">Its statement.</param>
/// <param name="Parent">The block it's inside, or null.</param>
public sealed record CodeBlock(string From, int Id, string Kind, string Text, int Line, int? Parent = null);

/// <summary>The page's data. See the top of this file.</summary>
/// <param name="Extra">Anything more the tool's own script wants. The views don't read it.</param>
public sealed record CodeModel(
  CodeSource Source,
  CodeLanguage Language,
  IReadOnlyList<CodeFile> Files,
  IReadOnlyList<CodeGroup> Groups,
  IReadOnlyList<CodeUnit> Units,
  IReadOnlyList<CodeMember> Members,
  IReadOnlyList<CodeEdge> UnitEdges,
  IReadOnlyList<CodeEdge> MemberEdges,
  IReadOnlyList<CodeClassEdge> ClassEdges,
  IReadOnlyList<CodeUnfollowed> Unfollowed,
  IReadOnlyList<CodeCall> Calls,
  IReadOnlyList<CodeBlock> Blocks,
  object? Extra = null)
{
  private static readonly JsonSerializerOptions JsonOptions = new()
  {
    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    // One property per line keeps git diffs of the page readable.
    WriteIndented = true,
    // The default encoder escapes <, > and &, so the JSON can't end the <script> element it sits in.
    Encoder = JavaScriptEncoder.Default,
  };

  /// <summary>
  /// Checks the model against its language and itself, so a mistake shows when the page is made: every kind, role,
  /// file, group and unit named is one that's listed.
  /// </summary>
  /// <exception cref="InvalidOperationException">Something names what isn't there.</exception>
  public void Validate()
  {
    Language.Validate();
    var files = Files.Select(f => f.Id).ToHashSet(StringComparer.Ordinal);
    var groups = Groups.Select(g => g.Id).ToHashSet(StringComparer.Ordinal);
    var units = Units.Select(u => u.Id).ToHashSet(StringComparer.Ordinal);
    var unitKinds = Language.UnitKinds.Select(k => k.Id).ToHashSet(StringComparer.Ordinal);
    var memberKinds = Language.MemberKinds.Select(k => k.Id).ToHashSet(StringComparer.Ordinal);
    var roles = Language.Roles.Select(r => r.Id).Append("none").ToHashSet(StringComparer.Ordinal);
    var relationships = Language.Relationships.Select(r => r.Id).ToHashSet(StringComparer.Ordinal);
    var classRelationships = Language.ClassRelationships.Select(r => r.Id).ToHashSet(StringComparer.Ordinal);

    foreach (var group in Groups.Where(g => g.Parent is not null))
    {
      Check(groups.Contains(group.Parent!), $"group '{group.Id}' is in '{group.Parent}', which isn't a group");
    }

    foreach (var unit in Units)
    {
      Check(unitKinds.Contains(unit.Kind), $"unit '{unit.Id}' is a '{unit.Kind}', which the language doesn't list");
      Check(files.Contains(unit.File), $"unit '{unit.Id}' is in file '{unit.File}', which isn't listed");
      Check(unit.Group is null || groups.Contains(unit.Group), $"unit '{unit.Id}' is in group '{unit.Group}', which isn't listed");
      foreach (var part in unit.Parts ?? [])
      {
        Check(files.Contains(part.File), $"unit '{unit.Id}' has a part in file '{part.File}', which isn't listed");
      }
    }

    foreach (var member in Members)
    {
      Check(units.Contains(member.Unit), $"member '{member.Id}' is in unit '{member.Unit}', which isn't listed");
      Check(memberKinds.Contains(member.Kind), $"member '{member.Id}' is a '{member.Kind}', which the language doesn't list");
      Check(roles.Contains(member.Role), $"member '{member.Id}' has the role '{member.Role}', which the language doesn't list");
      Check(files.Contains(member.File), $"member '{member.Id}' is in file '{member.File}', which isn't listed");
    }

    foreach (var edge in UnitEdges.Concat(MemberEdges))
    {
      Check(relationships.Contains(edge.Kind), $"an edge from '{edge.From}' is a '{edge.Kind}', which the language doesn't list");
    }

    foreach (var edge in ClassEdges)
    {
      Check(classRelationships.Contains(edge.Kind), $"a class diagram line from '{edge.From}' is a '{edge.Kind}', which the language doesn't list");
    }
  }

  private static void Check(bool holds, string problem)
  {
    if (!holds)
    {
      throw new InvalidOperationException($"The page's data is wrong: {problem}.");
    }
  }

  /// <summary>The model as the page's data.</summary>
  public string ToJson() => JsonSerializer.Serialize(this, JsonOptions);
}
