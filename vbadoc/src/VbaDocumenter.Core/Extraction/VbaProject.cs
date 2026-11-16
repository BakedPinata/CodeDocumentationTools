namespace VbaDocumenter.Core.Extraction;

/// <summary>The four kinds of VBA module. The kind decides the file extension and how the module is shown.</summary>
public enum VbaModuleKind
{
  /// <summary>Standard module (.bas).</summary>
  Standard,
  /// <summary>Class module (.cls).</summary>
  Class,
  /// <summary>UserForm (.frm).</summary>
  Form,
  /// <summary>ThisWorkbook and sheet modules (.cls).</summary>
  Document,
}

/// <summary>One module as stored in the workbook.</summary>
public sealed class VbaModule
{
  public required string Name { get; init; }
  public required VbaModuleKind Kind { get; init; }

  /// <summary>Name of the stream inside the VBA storage. Usually the same as <see cref="Name"/>.</summary>
  public required string StreamName { get; init; }

  /// <summary>
  /// The module's source exactly as stored (after decompression), in the project's code page.
  /// Kept as bytes so that writing it back out is lossless, even for characters that
  /// don't round-trip through .NET strings.
  /// </summary>
  public required byte[] SourceBytes { get; init; }

  /// <summary><see cref="SourceBytes"/> decoded with the project's code page.</summary>
  public required string Source { get; init; }

  /// <summary>
  /// For a module read from a folder of files: its file's path below that folder, with <c>/</c> between the
  /// parts, which the page refers to. Null for a module read from a workbook, whose file is the one vbadoc writes.
  /// </summary>
  public string? SourceFile { get; init; }

  /// <summary>For a module read from a folder: its file's bytes exactly as they are, which the page and parser use.</summary>
  public byte[]? FileBytes { get; init; }

  /// <summary>
  /// For a module read from a folder with <c>--folder-groups</c>: its subfolders, joined by dots as a <c>'@Folder</c>
  /// is, for a module that has no <c>'@Folder</c> of its own. Null otherwise.
  /// </summary>
  public string? FolderFromPath { get; init; }
}

/// <summary>What a library reference points at, which says how it can be found.</summary>
public enum VbaReferenceKind
{
  /// <summary>A type library registered on the machine.</summary>
  Registered,
  /// <summary>Another VBA project, such as an add-in, found by its path.</summary>
  Project,
  /// <summary>A library of controls, as forms use.</summary>
  Control,
}

/// <summary>A library reference (Tools > References in the VBA editor).</summary>
/// <param name="Name">The reference's name, such as "Scripting".</param>
/// <param name="Description">What the VBA editor calls it, such as "Microsoft Scripting Runtime".</param>
/// <param name="Guid">The GUID in its libid, with braces, or null if it has none.</param>
public sealed partial record VbaReference(string Name, string Description, string? Guid, VbaReferenceKind Kind)
{
  // Every project has these two, so there's nothing to bring or to tick: OLE Automation and the Office library.
  static readonly string[] EveryProjectHas = ["{00020430-0000-0000-C000-000000000046}", "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"];

  /// <summary>Its libid, which says where the library is.</summary>
  public string Libid { get; init; } = "";

  /// <summary>Its records in the project's dir stream, as they stand, so any kind can be copied to another project.</summary>
  public byte[] Records { get; init; } = [];

  /// <summary>Whether it's one Excel gives every project.</summary>
  public bool IsStandard => Guid is not null && EveryProjectHas.Contains(Guid, StringComparer.OrdinalIgnoreCase);

  /// <summary>The file a reference to another VBA project points at, or null for any other kind.</summary>
  public string? ProjectPath => Kind == VbaReferenceKind.Project && Libid.StartsWith(@"*\", StringComparison.Ordinal) && Libid.Length > 3 ? Libid[3..] : null;

  /// <summary>A reference from its name and libid, such as <c>*\G{…}#1.0#0#C:\…\scrrun.dll#Microsoft Scripting Runtime</c>.</summary>
  public static VbaReference From(string name, string libid, VbaReferenceKind kind)
  {
    var description = libid[(libid.LastIndexOf('#') + 1)..];
    var guid = GuidInLibid().Match(libid);
    return new VbaReference(name, description.Length > 0 ? description : name, guid.Success ? guid.Value.ToUpperInvariant() : null, kind) { Libid = libid };
  }

  [System.Text.RegularExpressions.GeneratedRegex(@"\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}")]
  private static partial System.Text.RegularExpressions.Regex GuidInLibid();
}

/// <summary>Everything read from a workbook's VBA project.</summary>
public sealed class VbaProject
{
  /// <summary>The VBA project name, e.g. "VBAProject".</summary>
  public required string Name { get; init; }

  /// <summary>Windows code page used for all module source, e.g. 1252.</summary>
  public required int CodePage { get; init; }

  /// <summary>Modules sorted by name (case-insensitive) so everything downstream is in a stable order.</summary>
  public required IReadOnlyList<VbaModule> Modules { get; init; }

  /// <summary>When the workbook was last saved, from its own document properties, or null when it doesn't say.</summary>
  public DateTimeOffset? Saved { get; init; }

  /// <summary>The project's library references, in order.</summary>
  public IReadOnlyList<VbaReference> References { get; init; } = [];

  /// <summary>
  /// Whether the project was read from a folder of module files rather than a workbook. Such a project has
  /// no references, as files don't record them, and nothing is written into its folder but the page.
  /// </summary>
  public bool FromFolder { get; init; }

  /// <summary>Modules whose name a file in another subfolder shares: documented, but not imported.</summary>
  public IReadOnlySet<string> ClashingModules { get; init; } = new HashSet<string>();
}
