using System.Text.RegularExpressions;
using VbaDocumenter.Core.Extraction;

namespace VbaDocumenter.Core.Output;

/// <summary>
/// A UserForm as the VBA editor's File > Export writes it, so its File > Import can bring it back: a <c>.frm</c> with the
/// designer's <c>Begin</c> block before the code, and a <c>.frx</c> beside it holding the form and its controls.
/// </summary>
public static partial class FormFiles
{
  /// <summary>The <c>.frx</c> beside a form's <c>.frm</c>, which the <c>.frm</c> names.</summary>
  public static string FrxNameFor(string frmName) => Path.ChangeExtension(frmName, ".frx");

  /// <summary>
  /// The <c>.frm</c>: <c>VERSION 5.00</c>, the <c>Begin</c> block naming the <c>.frx</c>, and the code with CRLF line
  /// endings and the attributes the VBA editor leaves out of an export (<c>VB_Base</c>, <c>VB_TemplateDerived</c> and
  /// <c>VB_Customizable</c>), in the project's code page.
  /// </summary>
  /// <param name="form">A form with its designer.</param>
  /// <param name="frmName">The <c>.frm</c>'s file name, whose <c>.frx</c> it names.</param>
  /// <param name="codePage">The project's code page.</param>
  public static byte[] Frm(VbaModule form, string frmName, int codePage)
  {
    ArgumentNullException.ThrowIfNull(form);
    var designer = form.Designer ?? throw new ArgumentException($"{form.Name} has no designer.", nameof(form));
    var code = ExportLeavesOut().Replace(form.Source.ReplaceLineEndings("\r\n"), "");
    return CodePages.Get(codePage).GetBytes(designer.ExportedFrame(Path.GetFileName(FrxNameFor(frmName))) + code);
  }

  [GeneratedRegex(@"^Attribute[ \t]+VB_(Base|TemplateDerived|Customizable)[ \t]*=[^\r\n]*\r\n", RegexOptions.IgnoreCase | RegexOptions.Multiline)]
  private static partial Regex ExportLeavesOut();
}
