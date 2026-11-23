using CodeDocumenter.Shared;
using System.Text;
using CodeDocumenter.Shared.Files;
using CodeDocumenter.Workbooks;
using VbaDocumenter.Core.Extraction;

namespace VbaDocumenter.Core.Output;

/// <summary>
/// Writes each module to a source file that the VBA editor can import, and keeps the output
/// folder in step with the workbook, as pqdoc keeps its .m files (<see cref="SourceFolderWriter"/>).
/// </summary>
public static class ProjectExporter
{
  /// <summary>
  /// The header the VBA editor writes at the top of an exported class. It isn't stored in the
  /// workbook, but without it the editor imports a .cls file as a standard module.
  /// </summary>
  public const string ClassHeader = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n";

  /// <summary>
  /// File types the tool owns: the modules' files, and a form's <c>.frx</c> (since step S4c). Any other file in the
  /// output folder is never touched.
  /// </summary>
  public static readonly IReadOnlyList<string> ModuleExtensions = [".bas", ".cls", ".frm", ".frx"];

  public const string GitAttributesContent =
      "# Created by vbadoc. Keep exported VBA exactly as the VBA editor expects it:\n" +
      "# CRLF line endings and the project's code page, with no conversion by git.\n" +
      "# (\"-text\" turns conversion off; \"diff\" keeps line-by-line diffs.)\n" +
      "*.bas -text diff\n" +
      "*.cls -text diff\n" +
      "*.frm -text diff\n" +
      "*.frx binary\n" +
      "*.html -text diff\n";

  /// <summary>The module's file: the one vbadoc writes for it, or for a module read from a folder, its own file's path.</summary>
  public static string FileNameFor(VbaModule module)
  {
    if (module.SourceFile is not null) return module.SourceFile;
    string extension = module.Kind switch
    {
      VbaModuleKind.Standard => ".bas",
      VbaModuleKind.Form => ".frm",
      _ => ".cls", // Class and Document
    };
    return SafeFileName(module.Name) + extension;
  }

  /// <summary>
  /// The exact bytes of the exported file: the module source with CRLF line endings, plus the class header where needed,
  /// and a UserForm's designer block as the VBA editor writes it (<see cref="FormFiles"/>), when it has a designer. For
  /// a module read from a folder, its own file's bytes, as they are.
  /// </summary>
  /// <param name="module">The module.</param>
  /// <param name="codePage">The project's code page, which a form's designer block is written in.</param>
  public static byte[] FileContent(VbaModule module, int codePage)
  {
    ArgumentNullException.ThrowIfNull(module);
    if (module.FileBytes is not null) return module.FileBytes;
    if (module.Kind == VbaModuleKind.Form && module.Designer is not null) return FormFiles.Frm(module, FileNameFor(module), codePage);
    var source = NormaliseLineEndings(module.SourceBytes);
    if (module.Kind is not (VbaModuleKind.Class or VbaModuleKind.Document)) return source;

    // The header is plain ASCII, which is the same in every code page VBA uses.
    return [.. Encoding.ASCII.GetBytes(ClassHeader), .. source];
  }

  /// <summary>
  /// A module's files, by name: its own, and for a UserForm with a designer, the <c>.frx</c> beside it, as the VBA
  /// editor's File > Export writes them, so its File > Import can bring the form back.
  /// </summary>
  public static IEnumerable<(string Name, byte[] Content)> Files(VbaModule module, int codePage)
  {
    ArgumentNullException.ThrowIfNull(module);
    var name = FileNameFor(module);
    yield return (name, FileContent(module, codePage));
    if (module.Kind == VbaModuleKind.Form && module.Designer is { } designer) yield return (FormFiles.FrxNameFor(name), designer.FrxBytes());
  }

  /// <summary>Every module's files, by name.</summary>
  /// <exception cref="ExtractionException">Two modules map to the same file name.</exception>
  public static Dictionary<string, byte[]> Files(IEnumerable<VbaModule> modules, int codePage)
  {
    ArgumentNullException.ThrowIfNull(modules);
    var files = new Dictionary<string, byte[]>(StringComparer.Ordinal);
    foreach (var (name, content) in modules.SelectMany(m => Files(m, codePage)))
    {
      if (files.Keys.Any(existing => existing.Equals(name, StringComparison.OrdinalIgnoreCase)))
        throw new ExtractionException($"Two modules map to the same file name '{name}'.");
      files[name] = content;
    }

    return files;
  }

  /// <summary>
  /// Writes the modules' files into the folder and keeps it in step with the workbook. A UserForm whose designer can't
  /// be read is written as its code alone, which the VBA editor can't import as a form: see
  /// <see cref="VbaProject.FormsWithoutDesigner"/>.
  /// </summary>
  public static SourceFolderResult Export(VbaProject project, string outputFolder)
  {
    ArgumentNullException.ThrowIfNull(project);

    // Work out every file first, so a problem (e.g. a name clash) stops us before anything changes on disk.
    var files = Files(project.Modules, project.CodePage);
    return SourceFolderWriter.Write(outputFolder, files, ModuleExtensions, GitAttributesContent);
  }

  /// <summary>Converts lone CR or LF bytes to CRLF. VBA source is normally CRLF already, so this is usually a no-op.</summary>
  internal static byte[] NormaliseLineEndings(byte[] source)
  {
    var result = new List<byte>(source.Length + 16);
    for (int i = 0; i < source.Length; i++)
    {
      byte b = source[i];
      if (b == '\r')
      {
        result.Add((byte)'\r');
        result.Add((byte)'\n');
        if (i + 1 < source.Length && source[i + 1] == '\n') i++;
      }
      else if (b == '\n')
      {
        result.Add((byte)'\r');
        result.Add((byte)'\n');
      }
      else
      {
        result.Add(b);
      }
    }
    return result.ToArray();
  }

  /// <summary>
  /// Replaces characters that Windows doesn't allow in file names. VBA module names are identifiers,
  /// so this should never change anything; it's a safety net. Uses a fixed list rather than
  /// Path.GetInvalidFileNameChars(), which differs between Windows and Linux.
  /// </summary>
  static string SafeFileName(string name)
  {
    const string invalid = "<>:\"/\\|?*";
    var sb = new StringBuilder(name.Length);
    foreach (char c in name) sb.Append(c < 32 || invalid.Contains(c) ? '_' : c);
    return sb.ToString();
  }
}
