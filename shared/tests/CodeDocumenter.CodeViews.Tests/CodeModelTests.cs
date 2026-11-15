using System.Text.Json;
using CodeDocumenter.CodeViews.BrowserTests;

namespace CodeDocumenter.CodeViews.Tests;

/// <summary>
/// The code model and a language's description of itself (step M3 of the generalisation plan), with the made-up
/// language the browser tests draw: what the page's data looks like, and that a mistake is caught when the page is
/// made.
/// </summary>
public sealed class CodeModelTests
{
  private static JsonElement Json(CodeModel model) => JsonDocument.Parse(model.ToJson()).RootElement;

  [Fact]
  public void The_data_has_the_models_parts_in_camelCase_and_leaves_out_what_isnt_set()
  {
    var data = Json(Toy.Model());

    Assert.Equal(
      ["source", "language", "files", "groups", "units", "members", "unitEdges", "memberEdges", "classEdges", "unfollowed", "calls", "blocks"],
      data.EnumerateObject().Select(p => p.Name));
    Assert.Equal("Toybox", data.GetProperty("source").GetProperty("title").GetString());
    Assert.Equal("embedded", data.GetProperty("source").GetProperty("sourceMode").GetString());

    var blueprint = data.GetProperty("units").EnumerateArray().Single(u => u.GetProperty("id").GetString() == "Blueprint");
    Assert.Equal(["id", "kind", "file", "usedBy"], blueprint.EnumerateObject().Select(p => p.Name));

    var language = data.GetProperty("language");
    Assert.Equal("gadgets", language.GetProperty("unit").GetProperty("many").GetString());
    Assert.Equal("overridden by", language.GetProperty("relationships")[1].GetProperty("dispatch").GetProperty("down").GetString());
  }

  [Fact]
  public void A_unit_can_be_in_two_files_and_a_file_can_hold_two_units()
  {
    var model = Toy.Model();

    var widget = model.Units.Single(u => u.Id == "Widget");
    Assert.Equal(["a.toy", "b.toy"], widget.Parts!.Select(p => p.File));
    Assert.Equal(["Widget", "Gizmo"], model.Units.Where(u => (u.Parts ?? []).Any(p => p.File == "b.toy")).Select(u => u.Id));
    Assert.Equal(["a.toy", "a.toy", "b.toy"], model.Members.Where(m => m.Unit == "Widget").Select(m => m.File));
  }

  [Fact]
  public void The_same_model_gives_the_same_data()
  {
    Assert.Equal(Toy.Model().ToJson(), Toy.Model().ToJson());
  }

  public static TheoryData<string, Func<CodeModel, CodeModel>> Mistakes => new()
  {
    { "unit 'Widget' is a 'gizmo'", m => m with { Units = [m.Units[0] with { Kind = "gizmo" }, .. m.Units.Skip(1)] } },
    { "unit 'Widget' is in file 'z.toy'", m => m with { Units = [m.Units[0] with { File = "z.toy" }, .. m.Units.Skip(1)] } },
    { "unit 'Widget' is in group 'Loft'", m => m with { Units = [m.Units[0] with { Group = "Loft" }, .. m.Units.Skip(1)] } },
    { "unit 'Widget' has a part in file 'z.toy'", m => m with { Units = [m.Units[0] with { Parts = [new("z.toy", 1, 2)] }, .. m.Units.Skip(1)] } },
    { "member 'Widget.Spin' is in unit 'Wodget'", m => m with { Members = [m.Members[0] with { Unit = "Wodget" }, .. m.Members.Skip(1)] } },
    { "member 'Widget.Spin' is a 'sub'", m => m with { Members = [m.Members[0] with { Kind = "sub" }, .. m.Members.Skip(1)] } },
    { "member 'Widget.Spin' has the role 'macro'", m => m with { Members = [m.Members[0] with { Role = "macro" }, .. m.Members.Skip(1)] } },
    { "an edge from 'Widget' is a 'implements'", m => m with { UnitEdges = [m.UnitEdges[0] with { Kind = "implements" }] } },
    { "a class diagram line from 'Gizmo' is a 'realization'", m => m with { ClassEdges = [m.ClassEdges[0] with { Kind = "realization" }] } },
    { "group 'Shop.Back' is in 'Loft'", m => m with { Groups = [m.Groups[0], m.Groups[1] with { Parent = "Loft" }] } },
  };

  [Theory]
  [MemberData(nameof(Mistakes))]
  public void A_model_that_names_what_isnt_listed_is_refused_with_what_and_where(string problem, Func<CodeModel, CodeModel> mistake)
  {
    var exception = Assert.Throws<InvalidOperationException>(() => mistake(Toy.Model()).Validate());

    Assert.StartsWith("The page's data is wrong: " + problem, exception.Message, StringComparison.Ordinal);
  }

  public static TheoryData<string, Func<CodeLanguage, CodeLanguage>> LanguageMistakes => new()
  {
    { "unit kind 'thing' is listed twice", l => l with { UnitKinds = [l.UnitKinds[0], l.UnitKinds[0]] } },
    { "unit kind 'a thing' isn't a plain name", l => l with { UnitKinds = [l.UnitKinds[0] with { Id = "a thing" }] } },
    { "'blob' isn't a shape the views draw", l => l with { UnitKinds = [l.UnitKinds[0] with { Shape = "blob" }] } },
    { "unit kind 'thing' has colour 9", l => l with { UnitKinds = [l.UnitKinds[0] with { Colour = 9 }] } },
    { "colour 2 has no label", l => l with { UnitKinds = [l.UnitKinds[0] with { Colour = 2 }] } },
    { "member kind 'step' is in category 'move'", l => l with { MemberKinds = [l.MemberKinds[0] with { Category = "move" }] } },
    { "exactly one relationship has to be the plain one", l => l with { Relationships = [l.Relationships[1]] } },
    { "relationship 'extends' has line 'wavy'", l => l with { Relationships = [l.Relationships[0], l.Relationships[1] with { Line = "wavy" }] } },
    { "relationship 'extends' needs its tags", l => l with { Relationships = [l.Relationships[0], l.Relationships[1] with { UsesTag = null }] } },
    { "class relationship 'inherit' has a choice of its own", l => l with { ClassRelationships = [l.ClassRelationships[0] with { Choice = "line:holds" }] } },
    { "role 'start' is an entry point of 'begin'", l => l with { Roles = [l.Roles[0] with { Entry = "begin" }] } },
  };

  [Theory]
  [MemberData(nameof(LanguageMistakes))]
  public void A_language_that_doesnt_hang_together_is_refused_with_why(string problem, Func<CodeLanguage, CodeLanguage> mistake)
  {
    var exception = Assert.Throws<InvalidOperationException>(() => mistake(Toy.Language).Validate());

    Assert.StartsWith("The language's description is wrong: " + problem, exception.Message, StringComparison.Ordinal);
  }
}
