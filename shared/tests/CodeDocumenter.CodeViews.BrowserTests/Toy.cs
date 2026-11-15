using System.Security.Cryptography;
using System.Text;
using CodeDocumenter.Shared;
using CodeDocumenter.Shared.Pages;

namespace CodeDocumenter.CodeViews.BrowserTests;

/// <summary>
/// A made-up language, Toy, and a small project in it, to show the views draw what a language describes and assume
/// nothing of VBA (step M3 of the generalisation plan). Toy has two kinds of unit (a thing and a plan), calls its
/// units gadgets and its members parts, and has a relationship of its own, Extends, which a call can follow. Its
/// project has a gadget in two files (Widget), a file with two gadgets (b.toy), and gadgets on shelves inside shelves.
/// </summary>
public static class Toy
{
  public const string Tool = "toydoc";

  public static CodeLanguage Language { get; } = new()
  {
    Name = "Toy",
    CodeBlock = "toy",
    Unit = new("gadget", "gadgets"),
    Member = new("part", "parts"),
    Callable = new("step", "steps"),
    GroupsLabel = "Shelves",
    GroupsTitle = "Gadgets on their shelves",
    UnitKinds =
    [
      new("thing", "Thing", "Things", CodeShapes.Square, 1, "thing"),
      new("plan", "Plan", "Plans", CodeShapes.Wye, 3, "plan", Hollow: true),
    ],
    ColourLabels = [new(1, "Thing"), new(3, "Plan")],
    MemberCategories = [new("step", "Steps", CodeShapes.Circle), new("fact", "Facts", CodeShapes.Square)],
    MemberKinds = [new("step", "Step", "step", Callable: true, HasBody: true), new("fact", "Fact", "fact")],
    Relationships =
    [
      new("uses", "Uses", "uses", "solid", "filled", 1, Plain: true),
      new("extends", "Extends", "extends", "dashed", "hollow", 2,
          UsesTag: "extends", UsedByTag: "extended by", UsesTitle: "Extends this", UsedByTitle: "Extends the selection",
          Dispatch: new("overridden by", "overrides")),
    ],
    ClassRelationships =
    [
      new("inherit", "Extends", "extends", "solid", "triangle", 2, "line:extends"),
      new("use", "Uses", "uses", "dashed", "open", 1, "line:uses"),
    ],
    Roles = [new("start", "Starting point", Entry: "start")],
    EntryGroups = [new("start", "Starting points")],
    ClassMarks = "+ open, - shut.",
    CallsNote = "The tree can't see steps wound up by hand.",
    Unfollowed = "loose",
    UnfollowedWhy = "Wound up by hand, so what it runs isn't written down",
    UnfollowedHeading = "Loose, so not followed:",
  };

  /// <summary>The project's files, by name.</summary>
  public static IReadOnlyDictionary<string, string> Sources { get; } = new Dictionary<string, string>
  {
    ["a.toy"] = string.Join("\n",
        "gadget Widget",
        "  step Spin",
        "    Gizmo.Run",
        "  end",
        "  fact Weight = 3",
        "end",
        ""),
    ["b.toy"] = string.Join("\n",
        "' more of Widget, and Gizmo",
        "",
        "more Widget",
        "  step Stop",
        "    Blueprint.Draw",
        "    thing.poke",
        "  end",
        "end",
        "",
        "gadget Gizmo extends Blueprint",
        "  step Run",
        "    ' nothing",
        "  end",
        "end",
        ""),
    ["c.toy"] = string.Join("\n",
        "plan Blueprint",
        "  step Draw",
        "    ' nothing",
        "  end",
        "end",
        ""),
  };

  public static CodeModel Model(bool embedSource = true)
  {
    var files = Sources.Select(s => new CodeFile(s.Key, s.Key,
        Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(s.Value)))[..16], s.Value.Split('\n').Length - 1)).ToList();
    var model = new CodeModel(
        new CodeSource("Toybox", embedSource ? "embedded" : "folder"),
        Language,
        files,
        [new CodeGroup("Shop", "Shop"), new CodeGroup("Shop.Back", "Back", "Shop")],
        [
          new CodeUnit("Widget", "thing", "a.toy", Group: "Shop.Back", Parts: [new("a.toy", 1, 6), new("b.toy", 3, 8)], Description: "Spins, and stops."),
          new CodeUnit("Gizmo", "thing", "b.toy", Group: "Shop", Parts: [new("b.toy", 10, 14)], UsedBy: 1),
          new CodeUnit("Blueprint", "plan", "c.toy", UsedBy: 2),
        ],
        [
          new CodeMember("Widget.Spin", "Widget", "Spin", "step", "public", "step Spin", "a.toy", 2, 4, 2, Role: "start", Class: new("+Spin", "")),
          new CodeMember("Widget.Weight", "Widget", "Weight", "fact", "private", "fact Weight = 3", "a.toy", 5, 5, 5, Unreferenced: true, Class: new("-Weight = 3")),
          new CodeMember("Widget.Stop", "Widget", "Stop", "step", "public", "step Stop", "b.toy", 4, 7, 4, Class: new("+Stop", "")),
          new CodeMember("Gizmo.Run", "Gizmo", "Run", "step", "public", "step Run", "b.toy", 11, 13, 11, Class: new("+Run", ""), UsedBy: 1),
          new CodeMember("Blueprint.Draw", "Blueprint", "Draw", "step", "public", "step Draw", "c.toy", 2, 4, 2, Class: new("+Draw", ""), UsedBy: 2),
        ],
        UnitEdges: [new("Widget", "Gizmo", "uses", 1), new("Widget", "Blueprint", "uses", 1), new("Gizmo", "Blueprint", "extends", 1)],
        MemberEdges:
        [
          new("Widget.Spin", "Gizmo.Run", "uses", 1, [3]),
          new("Widget.Stop", "Blueprint.Draw", "uses", 1, [5]),
          new("Gizmo.Run", "Blueprint.Draw", "extends", 1),
        ],
        ClassEdges: [new("Gizmo", "Blueprint", "inherit"), new("Widget", "Gizmo", "use"), new("Widget", "Blueprint", "use")],
        Unfollowed: [new("Widget.Stop", "thing.poke", 6)],
        Calls: [new("Widget.Spin", "Gizmo.Run", null, 3), new("Widget.Stop", "Blueprint.Draw", null, 5), new("Widget.Stop", null, "thing.poke", 6)],
        Blocks: []);
    model.Validate();
    return model;
  }

  /// <summary>The page a tool for Toy would make: the shell, the code views, and a few lines of its own.</summary>
  public static string Page(bool embedSource = true)
  {
    var sources = new StringBuilder();
    if (embedSource)
    {
      foreach (var (name, text) in Sources)
      {
        sources.Append($"<pre hidden class=\"{Tool}-source\" data-file=\"{PageBuilder.Escape(name)}\">\n").Append(PageBuilder.Escape(text)).Append("</pre>\n");
      }
    }

    return new ShellPage
    {
      Tool = Tool,
      Version = ToolVersion.Parse(Tool, "1.0.0+b7c0ebe8d952f156a36dd4ed8b545d572397f368"),
      PageTitle = "Toybox",
      Title = "Toybox",
      Tabs = CodeViewsPage.Tabs,
      SearchHtml = "    <label for=\"search\" class=\"visually-hidden\">Find a gadget or part</label>\n" +
          "    <input id=\"search\" type=\"search\" list=\"search-list\" placeholder=\"Find gadget or part…\">\n    <datalist id=\"search-list\"></datalist>\n",
      LoadFolderHint = "Choose the folder that holds this page's .toy files",
      Summary = "3 gadgets.",
      SourceWord = "toybox",
      PanelName = "Source",
      ExplorerLabel = "Gadgets and parts",
      ViewsHtml = CodeViewsPage.ViewsHtml,
      FaviconSvg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 16 16\"><rect width=\"16\" height=\"16\" fill=\"#888\"/></svg>",
      Stylesheet = CodeViewsPage.Stylesheet,
      DataJson = Model(embedSource).ToJson(),
      SourcesHtml = sources.ToString(),
      Scripts = CodeViewsPage.Scripts([new PageScript(typeof(Toy).Assembly, "Toy/toy.js")]),
    }.Render();
  }
}
