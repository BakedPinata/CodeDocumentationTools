using VbaDocumenter.Core.Extraction;
using VbaDocumenter.Core.Importing;

namespace VbaDocumenter.Tests.Importing;

/// <summary>A class exported by the VBA editor gets back the attributes a project stores and the export leaves out.</summary>
public sealed class ClassAttributesTests
{
  // What File > Export writes after its VERSION block, which ModuleFolderReader takes off.
  const string Exported =
    "Attribute VB_Name = \"Router\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n"
    + "Attribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\n'@Folder(\"Jira\")\r\nOption Explicit\r\n";

  // The same class as a project stores it, and as vbadoc exports it from a workbook.
  const string Stored =
    "Attribute VB_Name = \"Router\"\r\nAttribute VB_Base = \"0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}\"\r\n"
    + "Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = False\r\n"
    + "Attribute VB_Exposed = False\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = False\r\n"
    + "'@Folder(\"Jira\")\r\nOption Explicit\r\n";

  [Fact]
  public void An_exported_class_gets_the_three_attributes_the_export_left_out_in_Excels_order()
  {
    Assert.Equal(Stored, ClassAttributes.Complete(Exported));
  }

  [Fact]
  public void A_class_that_has_them_all_is_the_same_string()
  {
    Assert.Same(Stored, ClassAttributes.Complete(Stored));
    Assert.Null(ClassAttributes.Completion(Stored));
  }

  [Fact]
  public void Only_the_attribute_lines_at_the_top_are_replaced()
  {
    var change = ClassAttributes.Completion(Exported)!;

    Assert.Equal(Exported.IndexOf("'@Folder", StringComparison.Ordinal), change.Length);
    Assert.EndsWith("Attribute VB_Customizable = False\r\n", change.Header, StringComparison.Ordinal);
  }

  [Fact]
  public void A_class_keeps_its_own_values_and_spelling_and_attributes_Excel_doesnt_list_come_after()
  {
    var source = "attribute VB_Name = \"Factory\"\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n"
      + "Attribute VB_Description = \"Makes things\"\r\nOption Explicit\r\n";

    var completed = ClassAttributes.Complete(source);

    Assert.Equal(
      "attribute VB_Name = \"Factory\"\r\nAttribute VB_Base = \"0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}\"\r\n"
      + "Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\n"
      + "Attribute VB_Exposed = True\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = False\r\n"
      + "Attribute VB_Description = \"Makes things\"\r\nOption Explicit\r\n",
      completed);
  }

  [Fact]
  public void A_class_with_only_its_name_gets_them_all_with_the_line_breaks_it_uses()
  {
    var completed = ClassAttributes.Complete("Attribute VB_Name = \"ILogger\"\nOption Explicit\n");

    Assert.Equal(Stored.Replace("Router", "ILogger", StringComparison.Ordinal).Replace("'@Folder(\"Jira\")\r\n", "", StringComparison.Ordinal)
      .Replace("\r\n", "\n", StringComparison.Ordinal), completed);
  }

  [Fact]
  public void A_class_thats_only_its_name_line_still_gets_a_line_break_after_each()
  {
    Assert.Equal(8, ClassAttributes.Complete("Attribute VB_Name = \"Empty\"").Split("\r\n", StringSplitOptions.RemoveEmptyEntries).Length);
  }

  [Theory]
  [InlineData("Option Explicit\r\nAttribute VB_Name = \"Late\"\r\n")]
  [InlineData("")]
  public void A_source_that_doesnt_start_with_its_name_is_left_alone(string source)
  {
    Assert.Same(source, ClassAttributes.Complete(source));
  }

  [Theory]
  [InlineData(VbaModuleKind.Standard)]
  [InlineData(VbaModuleKind.Document)]
  public void Only_a_class_or_a_form_is_changed(VbaModuleKind kind)
  {
    var module = new VbaModule { Name = "Router", Kind = kind, StreamName = "Router", Source = Exported, SourceBytes = [] };

    Assert.Same(Exported, ClassAttributes.StoredSource(module));
  }

  // A form without VB_Base crashed Excel (step S4a). Its VB_Base names two GUIDs of its own, which Excel replaces when it
  // compiles, so any will do, and a form is predeclared.
  [Fact]
  public void An_exported_form_gets_a_VB_Base_of_its_own_and_the_other_two()
  {
    var form = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "UserForms", "frmDemoForm2.frm"));
    var exported = form[form.IndexOf("Attribute VB_Name", StringComparison.Ordinal)..];

    var lines = ClassAttributes.Complete(exported, VbaModuleKind.Form).Split("\r\n");

    Assert.Equal("Attribute VB_Name = \"frmDemoForm2\"", lines[0]);
    Assert.Matches(@"^Attribute VB_Base = ""0\{[0-9A-F-]{36}\}\{[0-9A-F-]{36}\}""$", lines[1]);
    Assert.Equal(
      ["Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = False", "Attribute VB_PredeclaredId = True",
       "Attribute VB_Exposed = False", "Attribute VB_TemplateDerived = False", "Attribute VB_Customizable = False"],
      lines[2..8]);
  }

  [Fact]
  public void Each_form_gets_a_new_VB_Base_and_a_form_that_has_one_keeps_it()
  {
    var exported = "Attribute VB_Name = \"Pick\"\r\nOption Explicit\r\n";
    var first = ClassAttributes.Complete(exported, VbaModuleKind.Form);

    Assert.NotEqual(first, ClassAttributes.Complete(exported, VbaModuleKind.Form));
    Assert.Same(first, ClassAttributes.Complete(first, VbaModuleKind.Form));
  }

  [Fact]
  public void Two_copies_of_a_form_compare_the_same_whatever_their_VB_Base()
  {
    VbaModule Form(string source) => new() { Name = "Pick", Kind = VbaModuleKind.Form, StreamName = "Pick", Source = source, SourceBytes = [] };
    var exported = "Attribute VB_Name = \"Pick\"\r\nOption Explicit\r\n";

    Assert.Equal(ClassAttributes.Comparable(Form(exported)), ClassAttributes.Comparable(Form(ClassAttributes.Complete(exported, VbaModuleKind.Form))));
    Assert.DoesNotContain("VB_Base", ClassAttributes.Comparable(Form(exported)), StringComparison.Ordinal);
  }
}
