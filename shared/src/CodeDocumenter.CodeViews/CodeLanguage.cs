namespace CodeDocumenter.CodeViews;

// What a language says about itself, so the views can draw its code without knowing the language: its kinds of unit,
// member and relationship, each with a label and how it's drawn, and its words for things. It travels in the page's
// data (CodeModel.Language), and the views take their legends, the Show… menu, the explorer's Types layout and their
// wording from it. Property names become camelCase in the JSON.

/// <summary>A word in the singular and the plural, lower case: <c>module</c>, <c>modules</c>.</summary>
public sealed record CodeWord(string One, string Many);

/// <summary>A kind of unit: what the graph's upper level draws as a node and the class diagram as a box.</summary>
/// <param name="Id">Its name in the data, such as <c>class</c>.</param>
/// <param name="Label">One of them: <c>Class</c>.</param>
/// <param name="Plural">Several: <c>Classes</c>, for the Show… menu and the explorer's Types layout.</param>
/// <param name="Shape">Its node's shape: one of <see cref="CodeShapes"/>.</param>
/// <param name="Colour">Its colour, as a slot of the views' palette: 1 to <see cref="CodeLanguage.Colours"/>.</param>
/// <param name="Stereotype">What its class diagram box says above its name, without the guillemets: <c>class</c>.</param>
/// <param name="Hollow">Whether its node is drawn as an outline, to tell it from another kind of the same shape and colour.</param>
public sealed record CodeUnitKind(string Id, string Label, string Plural, string Shape, int Colour, string Stereotype, bool Hollow = false);

/// <summary>What a colour stands for in the graph's legend at member level, where a member takes its unit's colour.</summary>
/// <param name="Slot">The palette's slot, as <see cref="CodeUnitKind.Colour"/>.</param>
/// <param name="Label">The kinds of unit that have it: <c>Class or interface</c>.</param>
public sealed record CodeColour(int Slot, string Label);

/// <summary>A group of kinds of member that the Show… menu chooses between and the graph gives one shape.</summary>
/// <param name="Id">Its name, such as <c>procedure</c>.</param>
/// <param name="Label">What the menu and the legend call it: <c>Procedures</c>.</param>
/// <param name="Shape">Its nodes' shape: one of <see cref="CodeShapes"/>.</param>
/// <param name="Under">
/// Another category it's shown with in the legend, having no entry of its own: an Enum's values go with Enums and
/// types.
/// </param>
public sealed record CodeMemberCategory(string Id, string Label, string Shape, string? Under = null);

/// <summary>A kind of member.</summary>
/// <param name="Id">Its name in the data, such as <c>propertyGet</c>.</param>
/// <param name="Label">One of them: <c>Property Get</c>.</param>
/// <param name="Category">The <see cref="CodeMemberCategory"/> it belongs to.</param>
/// <param name="Callable">Whether it can be called, so the Calls tab can start from it or show it.</param>
/// <param name="HasBody">Whether it has code of its own that makes calls, so a unit's Calls tree lists it.</param>
public sealed record CodeMemberKind(string Id, string Label, string Category, bool Callable = false, bool HasBody = false);

/// <summary>
/// A relationship that makes a call to one member run another: an interface's member and what implements it, an
/// event and what handles it. The Calls tab opens such a member to those others.
/// </summary>
/// <param name="Down">How the other is reached from the member called: <c>implemented by</c>.</param>
/// <param name="Up">The same, turned round: <c>implements</c>.</param>
/// <param name="FromRoles">Only edges from members with one of these roles count, or null for any.</param>
/// <param name="ToKinds">Only edges to members of one of these kinds count, or null for any.</param>
public sealed record CodeDispatch(string Down, string Up, IReadOnlyList<string>? FromRoles = null, IReadOnlyList<string>? ToKinds = null);

/// <summary>A kind of edge between units or members: a line in the graph.</summary>
/// <param name="Id">Its name in the data, such as <c>implements</c>.</param>
/// <param name="Label">What the legend and the Show… menu call it: <c>Implements</c>.</param>
/// <param name="Verb">What goes between two names to say it: <c>implements</c>, <c>handles or raises events of</c>.</param>
/// <param name="Line">How its line is drawn: <c>solid</c>, <c>dashed</c> or <c>dotted</c>.</param>
/// <param name="Arrow">Its arrowhead: <c>filled</c> or <c>hollow</c>.</param>
/// <param name="Colour">Its colour, as a slot of the views' line colours: 1 to <see cref="CodeLanguage.LineColours"/>.</param>
/// <param name="Plain">
/// True for the ordinary one (uses): the Dependencies list counts its references and gives it no tag.
/// </param>
/// <param name="UsesTag">Its tag on an entry of a Uses list: <c>implements</c>. Null when <paramref name="Plain"/>.</param>
/// <param name="UsedByTag">Its tag on an entry of a Used by list: <c>implemented by</c>.</param>
/// <param name="UsesTitle">The tag's tooltip in a Uses list.</param>
/// <param name="UsedByTitle">The tag's tooltip in a Used by list.</param>
/// <param name="ShowLabel">Its label in the Show… menu when that says more than <paramref name="Label"/>: <c>Events (graph)</c>.</param>
/// <param name="Dispatch">Set when a call can follow it (see <see cref="CodeDispatch"/>).</param>
public sealed record CodeRelationship(
  string Id, string Label, string Verb, string Line, string Arrow, int Colour, bool Plain = false,
  string? UsesTag = null, string? UsedByTag = null, string? UsesTitle = null, string? UsedByTitle = null,
  string? ShowLabel = null, CodeDispatch? Dispatch = null);

/// <summary>A kind of line in the class diagram.</summary>
/// <param name="Id">Its name in the data, such as <c>realization</c>.</param>
/// <param name="Label">What the legend calls it: <c>Implements</c>.</param>
/// <param name="Verb">What goes between two names to say it: <c>implements</c>.</param>
/// <param name="Line">How its line is drawn: <c>solid</c> or <c>dashed</c>.</param>
/// <param name="Arrow">Its arrowhead: <c>triangle</c> (hollow, closed) or <c>open</c>.</param>
/// <param name="Colour">A slot of the views' line colours, or 0 for the muted text colour.</param>
/// <param name="Choice">
/// The Show… choice that shows or hides it: <c>line:</c> and a <see cref="CodeRelationship"/>'s id, or one of its own,
/// which then needs <paramref name="ChoiceLabel"/>.
/// </param>
/// <param name="ChoiceLabel">The label of a choice of its own: <c>Holds in a variable (class diagram)</c>.</param>
public sealed record CodeClassRelationship(
  string Id, string Label, string Verb, string Line, string Arrow, int Colour, string Choice, string? ChoiceLabel = null);

/// <summary>What a member is for, beyond its kind: a macro, an event handler. Shown as a badge.</summary>
/// <param name="Id">Its name in the data, such as <c>macro</c>.</param>
/// <param name="Label">What the badge says: <c>Macro</c>.</param>
/// <param name="Entry">
/// The <see cref="CodeEntryGroup"/> its members are listed under on the Calls tab, when something outside the code
/// calls them; null otherwise.
/// </param>
public sealed record CodeRole(string Id, string Label, string? Entry = null);

/// <summary>A group of entry points the Calls tab starts from: <c>Macros</c>.</summary>
public sealed record CodeEntryGroup(string Id, string Label);

/// <summary>The shapes a node can have.</summary>
public static class CodeShapes
{
  public const string Circle = "circle";
  public const string Square = "square";
  public const string Diamond = "diamond";
  public const string Triangle = "triangle";
  public const string Star = "star";
  public const string Cross = "cross";
  public const string Wye = "wye";
}

/// <summary>A language, as the views need to know it. See the top of this file.</summary>
public sealed record CodeLanguage
{
  /// <summary>How many colours the views' palette has for kinds of unit.</summary>
  public const int Colours = 3;

  /// <summary>How many colours the views have for kinds of line.</summary>
  public const int LineColours = 3;

  /// <summary>The language's name, as people write it: <c>VBA</c>.</summary>
  public required string Name { get; init; }

  /// <summary>The name written on its code blocks in copied Markdown, such as <c>vba</c>.</summary>
  public required string CodeBlock { get; init; }

  /// <summary>What a unit is called: <c>module</c>.</summary>
  public required CodeWord Unit { get; init; }

  /// <summary>What a member is called: <c>member</c>.</summary>
  public required CodeWord Member { get; init; }

  /// <summary>What something that can be called is called: <c>procedure</c>.</summary>
  public required CodeWord Callable { get; init; }

  /// <summary>The explorer's layout of units in their groups: its label, such as <c>Folders</c>.</summary>
  public required string GroupsLabel { get; init; }

  /// <summary>That layout's tooltip: <c>Modules in their '@Folder folders</c>.</summary>
  public required string GroupsTitle { get; init; }

  /// <summary>In that layout, what the units in no group are listed under: <c>No folder</c>.</summary>
  public string Ungrouped { get; init; } = "No group";

  public required IReadOnlyList<CodeUnitKind> UnitKinds { get; init; }

  /// <summary>What each colour in use stands for, for the legend.</summary>
  public required IReadOnlyList<CodeColour> ColourLabels { get; init; }

  public required IReadOnlyList<CodeMemberCategory> MemberCategories { get; init; }

  public required IReadOnlyList<CodeMemberKind> MemberKinds { get; init; }

  /// <summary>The kinds of edge, in the order the legend lists them. Exactly one is <see cref="CodeRelationship.Plain"/>.</summary>
  public required IReadOnlyList<CodeRelationship> Relationships { get; init; }

  public required IReadOnlyList<CodeClassRelationship> ClassRelationships { get; init; }

  public IReadOnlyList<CodeRole> Roles { get; init; } = [];

  public IReadOnlyList<CodeEntryGroup> EntryGroups { get; init; } = [];

  /// <summary>Under the class diagram's legend, what the marks before a line mean: <c>+ public, − private, ~ friend.</c></summary>
  public string ClassMarks { get; init; } = string.Empty;

  /// <summary>
  /// On the Calls tab, what the tree can't see, completing "Calls the code makes directly, in its order.":
  /// <c>The tree can't see Application.Run…</c>
  /// </summary>
  public string CallsNote { get; init; } = string.Empty;

  /// <summary>What a call that can't be followed is called: <c>late-bound</c>.</summary>
  public string Unfollowed { get; init; } = "not followed";

  /// <summary>Why it can't: the tooltip beside one on the Calls tab.</summary>
  public string UnfollowedWhy { get; init; } = string.Empty;

  /// <summary>The heading over a member's calls that can't be followed, in the source panel: <c>Late-bound, so not resolved:</c></summary>
  public string UnfollowedHeading { get; init; } = "Not followed:";

  /// <summary>The tooltip on a member nothing refers to, saying what may use it all the same.</summary>
  public string UnreferencedWhy { get; init; } = "Nothing in the code refers to it.";

  /// <summary>
  /// Checks the description hangs together, so a mistake in a language shows when its page is made, not as a blank
  /// legend: ids are unique, every kind names a category, shapes and colours are ones the views have, and one
  /// relationship is the plain one.
  /// </summary>
  /// <exception cref="InvalidOperationException">It doesn't.</exception>
  public void Validate()
  {
    Unique(UnitKinds.Select(k => k.Id), "unit kind");
    Unique(MemberCategories.Select(c => c.Id), "member category");
    Unique(MemberKinds.Select(k => k.Id), "member kind");
    Unique(Relationships.Select(r => r.Id), "relationship");
    Unique(ClassRelationships.Select(r => r.Id), "class relationship");
    Unique(Roles.Select(r => r.Id), "role");
    Unique(EntryGroups.Select(g => g.Id), "entry group");

    string[] shapes = [CodeShapes.Circle, CodeShapes.Square, CodeShapes.Diamond, CodeShapes.Triangle, CodeShapes.Star, CodeShapes.Cross, CodeShapes.Wye];
    foreach (var shape in UnitKinds.Select(k => k.Shape).Concat(MemberCategories.Select(c => c.Shape)))
    {
      Check(shapes.Contains(shape), $"'{shape}' isn't a shape the views draw");
    }

    foreach (var kind in UnitKinds)
    {
      Check(kind.Colour is >= 1 and <= Colours, $"unit kind '{kind.Id}' has colour {kind.Colour}, and the palette has 1 to {Colours}");
      Check(ColourLabels.Any(c => c.Slot == kind.Colour), $"colour {kind.Colour} has no label");
    }

    var categories = MemberCategories.Select(c => c.Id).ToHashSet(StringComparer.Ordinal);
    foreach (var kind in MemberKinds)
    {
      Check(categories.Contains(kind.Category), $"member kind '{kind.Id}' is in category '{kind.Category}', which isn't listed");
    }

    foreach (var category in MemberCategories.Where(c => c.Under is not null))
    {
      Check(categories.Contains(category.Under!), $"member category '{category.Id}' goes under '{category.Under}', which isn't listed");
    }

    Check(Relationships.Count(r => r.Plain) == 1, "exactly one relationship has to be the plain one");
    foreach (var relationship in Relationships)
    {
      Check(relationship.Colour is >= 1 and <= LineColours, $"relationship '{relationship.Id}' has colour {relationship.Colour}, and there are 1 to {LineColours}");
      Check(relationship.Line is "solid" or "dashed" or "dotted", $"relationship '{relationship.Id}' has line '{relationship.Line}'");
      Check(relationship.Arrow is "filled" or "hollow", $"relationship '{relationship.Id}' has arrow '{relationship.Arrow}'");
      Check(relationship.Plain || (relationship.UsesTag is not null && relationship.UsedByTag is not null), $"relationship '{relationship.Id}' needs its tags");
    }

    foreach (var relationship in ClassRelationships)
    {
      Check(relationship.Colour is >= 0 and <= LineColours, $"class relationship '{relationship.Id}' has colour {relationship.Colour}");
      Check(relationship.Line is "solid" or "dashed", $"class relationship '{relationship.Id}' has line '{relationship.Line}'");
      Check(relationship.Arrow is "triangle" or "open", $"class relationship '{relationship.Id}' has arrow '{relationship.Arrow}'");
      var ownChoice = !Relationships.Any(r => "line:" + r.Id == relationship.Choice);
      Check(!ownChoice || relationship.ChoiceLabel is not null, $"class relationship '{relationship.Id}' has a choice of its own, which needs a label");
    }

    var entries = EntryGroups.Select(g => g.Id).ToHashSet(StringComparer.Ordinal);
    foreach (var role in Roles.Where(r => r.Entry is not null))
    {
      Check(entries.Contains(role.Entry!), $"role '{role.Id}' is an entry point of '{role.Entry}', which isn't listed");
    }
  }

  private static void Unique(IEnumerable<string> ids, string what)
  {
    var seen = new HashSet<string>(StringComparer.Ordinal);
    foreach (var id in ids)
    {
      // Ids go into class names and attributes unescaped, so they have to be plain.
      Check(id.Length > 0 && id.All(c => char.IsAsciiLetterOrDigit(c) || c == '-'), $"{what} '{id}' isn't a plain name");
      Check(seen.Add(id), $"{what} '{id}' is listed twice");
    }
  }

  private static void Check(bool holds, string problem)
  {
    if (!holds)
    {
      throw new InvalidOperationException($"The language's description is wrong: {problem}.");
    }
  }
}
