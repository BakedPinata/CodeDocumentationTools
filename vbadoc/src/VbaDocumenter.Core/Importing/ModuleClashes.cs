using CodeDocumenter.Workbooks.Importing;
using VbaDocumenter.Core.Analysis;
using VbaDocumenter.Core.Extraction;
using VbaDocumenter.Core.Parsing;

namespace VbaDocumenter.Core.Importing;

/// <summary>
/// Finds what's in the way of a choice of modules, and says what's at stake with each: what else in the target uses
/// what would be replaced, and what's coming that would use the target's version instead.
/// </summary>
public static class ModuleClashes
{
  /// <summary>Compares the plan's modules with the target's.</summary>
  /// <param name="source">The project they come from.</param>
  /// <param name="sourceName">The source's file name.</param>
  /// <param name="plan">What's coming.</param>
  /// <param name="import">What uses what in the source.</param>
  /// <param name="target">The target.</param>
  public static TargetComparison Compare(VbaProject source, string sourceName, ImportPlan plan, ModuleImport import, TargetWorkbook target)
  {
    ArgumentNullException.ThrowIfNull(source);
    ArgumentNullException.ThrowIfNull(plan);
    ArgumentNullException.ThrowIfNull(import);
    ArgumentNullException.ThrowIfNull(target);

    var incoming = source.Modules.ToDictionary(m => m.Name, StringComparer.OrdinalIgnoreCase);
    var theirs = (target.Project?.Modules ?? []).ToDictionary(m => m.Name, StringComparer.OrdinalIgnoreCase);
    var codeNames = target.CodeNames.ToHashSet(StringComparer.OrdinalIgnoreCase);
    var usedInTarget = new Lazy<ILookup<string, string>>(() => UsedBy(target.Project));

    var added = new List<string>();
    var identical = new List<string>();
    var clashes = new List<ImportClash>();
    foreach (var id in plan.Items)
    {
      var module = incoming[id];
      var comingUsers = plan.Items.Where(item => !item.Equals(id, StringComparison.OrdinalIgnoreCase) && import.Uses(item).Contains(id, StringComparer.OrdinalIgnoreCase)).ToList();
      if (codeNames.Contains(id))
      {
        clashes.Add(CantBeReplaced(id, "a sheet's or the workbook's module", comingUsers));
      }
      else if (!theirs.TryGetValue(id, out var existing))
      {
        added.Add(id);
      }
      else if (existing.Kind == module.Kind && Same(existing, module) && SameDesigner(existing, module))
      {
        identical.Add(id);
      }
      else if (existing.Kind is VbaModuleKind.Document or VbaModuleKind.Form)
      {
        clashes.Add(CantBeReplaced(id, existing.Kind == VbaModuleKind.Form ? "a UserForm" : "a sheet's or the workbook's module", comingUsers));
      }
      else
      {
        var problem = existing.Kind == module.Kind
          ? $"The target already has {id}, and it differs from the one in {sourceName}."
          : $"The target's {id} is {KindOf(existing.Kind)}, and the one in {sourceName} is {KindOf(module.Kind)}.";
        clashes.Add(new ImportClash(id, problem, ImportClash.Overwriting(usedInTarget.Value[id].ToList()), ImportClash.Skipping(comingUsers)));
      }
    }

    return new TargetComparison(added, identical, clashes);
  }

  // A form is the same only if its designer is: the same code with other controls is another form.
  static bool SameDesigner(VbaModule a, VbaModule b) =>
    a.Kind != VbaModuleKind.Form || a.Designer is not null && b.Designer is not null && a.Designer.SameAs(b.Designer);

  // Only what's at the end can differ without meaning anything: the VBA editor keeps or drops a last line as it likes.
  // They're compared as a project stores them, so a class exported by the VBA editor matches the same class in a workbook,
  // and a form without its VB_Base, which each copy of a form has its own of.
  static bool Same(VbaModule a, VbaModule b) =>
    ClassAttributes.Comparable(a).TrimEnd().Equals(ClassAttributes.Comparable(b).TrimEnd(), StringComparison.Ordinal);

  static ILookup<string, string> UsedBy(VbaProject? project) =>
    (project is null ? [] : DependencyAnalyser.Analyse(ProjectParser.Parse(project)).ModuleEdges)
      .Where(e => !e.From.Equals(e.To, StringComparison.OrdinalIgnoreCase))
      .Select(e => (e.To, e.From)).Distinct()
      .ToLookup(e => e.To, e => e.From, StringComparer.OrdinalIgnoreCase);

  static ImportClash CantBeReplaced(string id, string what, List<string> comingUsers) => new(
    id,
    $"{id} is {what} in the target, which can't be replaced.",
    null,
    comingUsers.Count == 0 ? "Leaves this one out." : $"Leaves this one out, so {ImportClash.Names(comingUsers)} will find the target's {id} instead.");

  static string KindOf(VbaModuleKind kind) => kind switch
  {
    VbaModuleKind.Standard => "a standard module",
    VbaModuleKind.Class => "a class",
    VbaModuleKind.Form => "a UserForm",
    _ => "a sheet's or the workbook's module",
  };
}
