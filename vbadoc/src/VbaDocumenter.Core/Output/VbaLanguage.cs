using CodeDocumenter.CodeViews;

namespace VbaDocumenter.Core.Output;

/// <summary>
/// VBA, as the code views library needs to know it (<see cref="CodeLanguage"/>): its kinds of module and member, the
/// lines between them, what calls a procedure from outside the code, and its words for things. The ids are the names
/// of the parser's and analyser's enums in camelCase, which is how <see cref="PageData"/> writes them.
/// </summary>
public static class VbaLanguage
{
  /// <summary>The Show… choice of the class diagram's own line, a variable that holds another module.</summary>
  public const string HoldsChoice = "line:holds";

  public static CodeLanguage Description { get; } = new()
  {
    Name = "VBA",
    CodeBlock = "vba",
    Unit = new("module", "modules"),
    Member = new("member", "members"),
    Callable = new("procedure", "procedures"),
    GroupsLabel = "Folders",
    GroupsTitle = "Modules in their '@Folder folders",
    Ungrouped = "No folder",

    // Colour carries three groups (validated as a palette); shape tells the five kinds apart.
    UnitKinds =
    [
      new("standard", "Standard module", "Standard modules", CodeShapes.Square, 1, "module"),
      new("class", "Class", "Classes", CodeShapes.Circle, 2, "class"),
      new("interface", "Interface", "Interfaces", CodeShapes.Circle, 2, "interface", Hollow: true),
      new("form", "UserForm", "UserForms", CodeShapes.Diamond, 3, "form"),
      new("document", "Document module", "Document modules", CodeShapes.Triangle, 3, "document"),
    ],
    ColourLabels = [new(1, "Standard module"), new(2, "Class or interface"), new(3, "UserForm or document")],

    MemberCategories =
    [
      new("procedure", "Procedures", CodeShapes.Circle),
      new("property", "Properties", CodeShapes.Diamond),
      new("data", "Variables and constants", CodeShapes.Square),
      new("type", "Enums and types", CodeShapes.Triangle),
      new("enumValue", "Enum values", CodeShapes.Triangle, Under: "type"),
      new("event", "Events", CodeShapes.Star),
      new("declare", "Declares", CodeShapes.Cross),
    ],
    MemberKinds =
    [
      new("sub", "Sub", "procedure", Callable: true, HasBody: true),
      new("function", "Function", "procedure", Callable: true, HasBody: true),
      new("propertyGet", "Property Get", "property", Callable: true, HasBody: true),
      new("propertyLet", "Property Let", "property", Callable: true, HasBody: true),
      new("propertySet", "Property Set", "property", Callable: true, HasBody: true),
      new("variable", "Variable", "data"),
      new("constant", "Const", "data"),
      new("enum", "Enum", "type"),
      new("enumMember", "Enum value", "enumValue"),
      new("type", "Type", "type"),
      new("event", "Event", "event", Callable: true),
      new("declare", "Declare", "declare", Callable: true),
    ],

    Relationships =
    [
      new("uses", "Uses", "uses", "solid", "filled", 1, Plain: true),
      new("implements", "Implements", "implements", "dashed", "hollow", 2,
          UsesTag: "implements", UsedByTag: "implemented by",
          UsesTitle: "Implements this interface or interface member", UsedByTitle: "Implements the selection",
          Dispatch: new("implemented by", "implements")),
      new("events", "Events", "handles or raises events of", "dotted", "filled", 3,
          UsesTag: "events", UsedByTag: "events",
          UsesTitle: "Raises or handles this event, or holds a WithEvents variable of this class",
          UsedByTitle: "Raises or handles the selection's events",
          ShowLabel: "Events (graph)",
          Dispatch: new("handled by", "handles", FromRoles: ["withEventsHandler"], ToKinds: ["event"])),
    ],
    ClassRelationships =
    [
      new("realization", "Implements", "implements", "dashed", "triangle", 2, "line:implements"),
      new("association", "Holds in a variable (labelled)", "holds a", "solid", "open", 0, HoldsChoice, "Holds in a variable (class diagram)"),
      new("dependency", "Uses", "uses", "dashed", "open", 1, "line:uses"),
    ],

    Roles =
    [
      new("macro", "Macro", Entry: "macro"),
      new("worksheetFunction", "Possible worksheet function", Entry: "worksheetFunction"),
      new("builtInEventHandler", "Event handler", Entry: "eventHandler"),
      new("withEventsHandler", "WithEvents handler"),
      new("controlEventHandler", "Control event handler", Entry: "eventHandler"),
      new("interfaceImplementation", "Implements an interface member"),
    ],
    EntryGroups = [new("macro", "Macros"), new("eventHandler", "Event handlers"), new("worksheetFunction", "Worksheet functions")],

    ClassMarks = "+ public, − private, ~ friend.",
    CallsNote = "The tree can't see Application.Run, OnAction or other calls made by name, and a late-bound call shows only its expression.",
    Unfollowed = "late-bound",
    UnfollowedWhy = "Called on an Object or Variant, so what it calls can't be known without running the code",
    UnfollowedHeading = "Late-bound, so not resolved:",
    UnreferencedWhy = "Nothing in the code refers to it. It may still be used by Application.Run, a button or late-bound code.",
  };
}
