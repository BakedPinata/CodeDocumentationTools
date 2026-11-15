using System.Security.Cryptography;
using System.Text.Json;
using CodeDocumenter.CodeViews;
using CodeDocumenter.Workbooks.Importing;
using VbaDocumenter.Core.Analysis;
using VbaDocumenter.Core.Extraction;
using VbaDocumenter.Core.Importing;
using VbaDocumenter.Core.Parsing;

namespace VbaDocumenter.Core.Output;

/// <summary>What vbadoc's own script wants of the project, beyond the code model: how module files are decoded, and
/// what the Export tab offers.</summary>
/// <param name="Name">The VBA project's name.</param>
/// <param name="Encoding">The browser's name for <paramref name="CodePage"/>.</param>
/// <param name="References">The project's library references, apart from the two every project has, for the Export tab.</param>
public sealed record PageProject(string Name, int CodePage, string Encoding, IReadOnlyList<PageReference> References);

/// <summary>A library reference (see <see cref="VbaReference"/>).</summary>
public sealed record PageReference(string Name, string Description);

/// <summary>What a module relies on that doesn't come with it (see <see cref="ImportNeed"/>).</summary>
public sealed record PageNeed(string What, string Why);

/// <summary>What vbadoc's own script wants of a module, for the Export tab.</summary>
/// <param name="WhyNot">Why the Export tab can't pick it, or null when it can (see <see cref="ModuleImport"/>).</param>
/// <param name="Needs">What it relies on that doesn't come with it, or null for nothing.</param>
public sealed record PageModuleExport(string? WhyNot, IReadOnlyList<PageNeed>? Needs);

/// <summary>
/// The page's data: a VBA project as the code model the code views library draws (<see cref="CodeModel"/>), in
/// <see cref="VbaLanguage"/>'s terms. A module is a unit and the whole of its exported file; a '@Folder is a group.
/// </summary>
public static class PageData
{
  /// <summary>Why a handler on a sheet has no control to show.</summary>
  const string ControlOnSheet = "the control is on the sheet, not in the VBA project";

  public static CodeModel From(DependencyModel model, VbaProject project, string title, bool embedSource)
  {
    var filesByName = project.Modules.ToDictionary(m => m.Name, StringComparer.Ordinal);
    string Camel(Enum value) => JsonNamingPolicy.CamelCase.ConvertName(value.ToString());
    var usedBy = Usage.UsedByOtherModules(model);
    int? UsedBy(string id) => usedBy.TryGetValue(id, out var count) ? count : null;
    var import = new ModuleImport(model);
    string FileOf(string module) => ProjectExporter.FileNameFor(filesByName[module]);

    var files = model.Modules.Select(node =>
    {
      var content = ProjectExporter.FileContent(filesByName[node.Id]);
      return new CodeFile(FileOf(node.Id), FileOf(node.Id), Convert.ToHexStringLower(SHA256.HashData(content))[..16], node.Module.LineCount);
    }).ToList();

    // A '@Folder("Import.Parsing") is the group Parsing inside the group Import.
    var groups = new List<CodeGroup>();
    var groupIds = new HashSet<string>(StringComparer.Ordinal);
    string? GroupOf(string? folder)
    {
      var parts = (folder ?? "").Split('.').Select(p => p.Trim()).Where(p => p.Length > 0).ToList();
      string? parent = null;
      foreach (var part in parts)
      {
        var id = parent is null ? part : parent + "." + part;
        if (groupIds.Add(id)) groups.Add(new CodeGroup(id, part, parent));
        parent = id;
      }
      return parent;
    }

    var units = model.Modules.Select(node =>
    {
      var needs = import.Needs(node.Id).Select(n => new PageNeed(n.What, n.Why)).ToList();
      var predeclared = node.Module.Kind == VbaModuleKind.Class && !node.IsInterface && node.Module.HasPredeclaredId;
      var whyNot = import.WhyNot(node.Id);
      return new CodeUnit(node.Id, node.IsInterface ? "interface" : Camel(node.Module.Kind), FileOf(node.Id),
          Group: GroupOf(node.Module.Folder),
          KindLabel: predeclared ? "Predeclared class" : null,
          Note: predeclared ? "Predeclared" : null,
          UsedBy: UsedBy(node.Id),
          Description: node.Module.Description,
          Extra: whyNot is null && needs.Count == 0 ? null : new PageModuleExport(whyNot, needs.Count > 0 ? needs : null));
    }).ToList();

    var classLines = model.Modules.ToDictionary(m => m.Id, m => ClassDiagram.Lines(m.Module));
    var members = model.Members.Where(n => n.Member != null).Select(n => new CodeMember(
        n.Id, n.Module.Name, n.Member!.Name, Camel(n.Member.Kind), Camel(n.Member.Visibility),
        MemberFormatter.Signature(n.Member), FileOf(n.Module.Name), n.Member.StartLine, n.Member.EndLine, n.Member.DeclarationLine,
        Camel(n.Role),
        n.Role == MemberRole.ControlEventHandler && n.Module.Kind == VbaModuleKind.Document ? ControlOnSheet : null,
        n.IsUnreferenced, n.Member.ParentName,
        classLines[n.Module.Name].GetValueOrDefault(n.Member), UsedBy(n.Id), n.Member.Description)).ToList();

    CodeEdge Edge(Edge e) => new(e.From, e.To, Camel(e.Kind), e.Count, e.Lines.Count == 0 ? null : e.Lines);
    var callers = model.Calls.Select(c => c.From).ToHashSet(StringComparer.Ordinal);

    var data = new CodeModel(
        new CodeSource(title, embedSource ? "embedded" : "folder"),
        VbaLanguage.Description,
        files,
        groups,
        units,
        members,
        model.ModuleEdges.Select(Edge).ToList(),
        model.MemberEdges.Select(Edge).ToList(),
        ClassDiagram.Edges(model),
        model.LateBound.Select(l => new CodeUnfollowed(l.FromId, l.Expression, l.Line)).ToList(),
        model.Calls.Select(c => new CodeCall(c.From, c.To, c.LateBound, c.Line, c.Block, c.Condition)).ToList(),
        // Only blocks with a call in them are drawn, so only procedures with calls need theirs.
        model.Members.Where(n => n.Member is { Blocks.Count: > 0 } && callers.Contains(n.Id))
            .SelectMany(n => n.Member!.Blocks.Select(b => new CodeDocumenter.CodeViews.CodeBlock(n.Id, b.Id, b.Kind, b.Text, b.Line, b.Parent))).ToList(),
        new PageProject(project.Name, project.CodePage, CodePages.Get(project.CodePage).WebName,
            project.References.Where(r => !r.IsStandard).Select(r => new PageReference(r.Name, r.Description)).ToList()));
    data.Validate();
    return data;
  }
}
