using System.Text;
using VbaDocumenter.Core.Extraction;

namespace VbaDocumenter.Core.Importing;

/// <summary>
/// The MSForms library, which a UserForm's code needs a reference to as soon as it names one of its types, as an event
/// such as <c>KeyPress</c> does (<c>MSForms.ReturnInteger</c>). A project the VBA editor gives a form refers to it
/// through a cache made on each machine; a reference to <c>FM20.DLL</c> by its GUID works as well, and Excel keeps it as
/// it is (step S4a). So forms going into a project with no MSForms reference bring that one.
/// </summary>
public static class FormsLibrary
{
  /// <summary>The library's GUID, which a project's control reference to it has as its original library too.</summary>
  public const string Guid = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}";

  const string Libid = @"*\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#C:\Windows\System32\FM20.DLL#Microsoft Forms 2.0 Object Library";

  /// <summary>The reference forms bring, registered, as the VBA editor shows it: MSForms.</summary>
  public static VbaReference Reference { get; } = VbaReference.From("MSForms", Libid, VbaReferenceKind.Registered) with { Records = Records() };

  /// <summary>
  /// Whether forms being added need the reference: one of them is a form, and neither the target nor the references
  /// being added have the library, known by its GUID. A target with another library called MSForms is left alone.
  /// </summary>
  public static bool IsNeeded(IEnumerable<VbaModule> adding, VbaProject? target, IEnumerable<VbaReference> added)
  {
    ArgumentNullException.ThrowIfNull(adding);
    ArgumentNullException.ThrowIfNull(added);
    if (!adding.Any(m => m.Kind == VbaModuleKind.Form)) return false;
    var references = (target?.References ?? []).Concat(added).ToList();
    return !references.Any(r => Guid.Equals(r.Guid, StringComparison.OrdinalIgnoreCase) || r.Name.Equals(Reference.Name, StringComparison.OrdinalIgnoreCase));
  }

  // REFERENCENAME, in the code page and in UTF-16, and REFERENCEREGISTERED (MS-OVBA 2.3.4.2.2). The libid is ASCII, the
  // same in every code page VBA uses.
  static byte[] Records()
  {
    var libid = Encoding.ASCII.GetBytes(Libid);
    return DirRecords.Write(
    [
      new(0x0016, Encoding.ASCII.GetBytes("MSForms")),
      new(0x003E, Encoding.Unicode.GetBytes("MSForms")),
      new(0x000D, [.. BitConverter.GetBytes(libid.Length), .. libid, 0, 0, 0, 0, 0, 0]),
    ]);
  }
}
