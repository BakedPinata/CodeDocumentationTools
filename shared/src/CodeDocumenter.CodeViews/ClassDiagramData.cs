namespace CodeDocumenter.CodeViews;

// The class diagram's part of the page's data, which a language's Core fills in and classes.js draws. Property
// names become camelCase in the JSON.

/// <summary>
/// One line in a class diagram box, as plain text, which the page draws as it is.
/// A line with <see cref="Parameters"/> (possibly empty) is drawn as a method, e.g. "+Describe(x As Long) : String";
/// one without is drawn as an attribute, e.g. "-m_Id : String". Attributes and methods go in separate compartments.
/// </summary>
/// <param name="Text">Visibility and name, plus the type for attributes: "+Describe", "-Const SEP : String = ",""</param>
/// <param name="Parameters">For methods, the parameter list without the brackets.</param>
/// <param name="Type">For methods, the return type.</param>
public sealed record PageClassLine(string Text, string? Parameters = null, string? Type = null);

/// <summary>A relationship between two boxes in the class diagram.</summary>
/// <param name="Kind">"realization" (From implements To), "association" (From holds a To in a module-level
/// variable) or "dependency" (any other use of To).</param>
/// <param name="Label">For associations, the variable name ("WithEvents m_svc" for event sources).</param>
public sealed record PageClassEdge(string From, string To, string Kind, string? Label);
