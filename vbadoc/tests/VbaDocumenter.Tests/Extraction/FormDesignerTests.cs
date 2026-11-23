using VbaDocumenter.Core.Extraction;

namespace VbaDocumenter.Tests.Extraction;

/// <summary>A UserForm's designer, read from the VBA editor's <c>.frm</c> and <c>.frx</c> (step S4b).</summary>
public sealed class FormDesignerTests
{
  static readonly string Folder = Path.Combine(AppContext.BaseDirectory, "Fixtures", "UserForms");

  static string Frm(string name) => File.ReadAllText(Path.Combine(Folder, name + ".frm"));

  static byte[] Frx(string name) => File.ReadAllBytes(Path.Combine(Folder, name + ".frx"));

  static (FormDesigner? Designer, string? Problem) Read(string name, byte[]? frx = null) =>
    FormDesigner.FromFiles(Frm(name), file => file == name + ".frx" ? frx ?? Frx(name) : null);

  [Fact]
  public void The_frame_is_the_Begin_block_without_the_line_naming_the_frx()
  {
    var designer = Read("frmDemoForm1").Designer!;

    Assert.Equal(
      "VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDemoForm1 \r\n   Caption         =   \"UserForm1\"\r\n"
      + "   ClientHeight    =   4304\r\n   ClientLeft      =   96\r\n   ClientTop       =   416\r\n   ClientWidth     =   5392\r\n"
      + "   StartUpPosition =   1  'CenterOwner\r\nEnd\r\n",
      designer.Frame);
  }

  [Fact]
  public void The_storage_is_the_frx_after_its_header_and_the_frx_is_kept_as_it_was()
  {
    var designer = Read("frmDemoForm2").Designer!;

    Assert.Equal(Frx("frmDemoForm2")[FormDesigner.FrxHeaderLength..], designer.Storage);
    Assert.Equal(Frx("frmDemoForm2"), designer.FrxBytes());
  }

  // The header's last two numbers, which a .frx vbadoc builds has to give, look like the form's outer size (step S4a).
  [Theory]
  [InlineData("frmDemoForm1")]
  [InlineData("frmDemoForm2")]
  public void A_frx_built_from_the_designer_has_the_header_the_VBA_editor_wrote(string name)
  {
    var read = Read(name).Designer!;
    var built = new FormDesigner(read.Frame, read.Storage).FrxBytes();

    Assert.Equal(Frx(name), built);
  }

  [Fact]
  public void The_Begin_block_comes_back_as_the_VBA_editor_exported_it()
  {
    var frm = Frm("frmDemoForm1");

    var exported = Read("frmDemoForm1").Designer!.ExportedFrame("frmDemoForm1.frx");

    Assert.Equal(frm[..frm.IndexOf("Attribute VB_Name", StringComparison.Ordinal)], exported);
  }

  [Fact]
  public void A_designer_is_the_same_as_itself_and_not_as_another_form()
  {
    var one = Read("frmDemoForm1").Designer!;

    Assert.True(one.SameAs(Read("frmDemoForm1").Designer!));
    Assert.False(one.SameAs(Read("frmDemoForm2").Designer!));
    Assert.True(one.SameAs(one with { Frame = one.Frame.Replace("End\r\n", "   TypeInfoVer     =   8\r\nEnd\r\n", StringComparison.Ordinal) }));
  }

  [Fact]
  public void A_frx_that_isnt_one_says_so()
  {
    var frx = Frx("frmDemoForm1");

    Assert.Equal("frmDemoForm1.frx isn't a .frx the VBA editor wrote", Read("frmDemoForm1", [.. "XX"u8, .. frx[2..]]).Problem);
    Assert.Equal("frmDemoForm1.frx is cut short", Read("frmDemoForm1", frx[..1000]).Problem);
    Assert.StartsWith("frmDemoForm1.frx can't be read", Read("frmDemoForm1", [.. frx[..24], .. new byte[frx.Length - 24]]).Problem, StringComparison.Ordinal);
  }

  [Fact]
  public void A_Begin_block_that_names_no_frx_says_so()
  {
    var frm = string.Join("\r\n", Frm("frmDemoForm1").Split("\r\n").Where(l => !l.Contains("OleObjectBlob", StringComparison.Ordinal)));

    Assert.Equal("Its .frm's Begin block doesn't name its .frx (no OleObjectBlob line)", FormDesigner.FromFiles(frm, _ => Frx("frmDemoForm1")).Problem);
  }
}
