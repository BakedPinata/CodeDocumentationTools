using System.Text;
using System.Text.RegularExpressions;
using CodeDocumenter.Shared.Files;
using CodeDocumenter.Shared;

namespace VbaDocumenter.Core.Extraction;

/// <summary>What reading a folder found to warn about, for the log.</summary>
/// <param name="Project">The modules, as a project with no references.</param>
/// <param name="Warnings">Each thing worth saying: a name that differs from its file's, a clash, files left out.</param>
public sealed record ModuleFolder(VbaProject Project, IReadOnlyList<string> Warnings);

/// <summary>
/// Reads a folder of exported VBA modules, as the VBA editor, Rubberduck or vbadoc wrote them or as they came from
/// elsewhere: <c>.bas</c>, <c>.cls</c> and <c>.frm</c> files, at the top level or, with
/// <c>recursive</c>, below it too. Nothing is written. Each module keeps its file's path and bytes, so the page refers
/// to the file where it is and its lines are the file's.
/// </summary>
public static partial class ModuleFolderReader
{
  /// <summary>The files read; any other file is left alone. A form's <c>.frx</c> is beside its <c>.frm</c>.</summary>
  public static IReadOnlyList<string> Extensions { get; } = [".bas", ".cls", ".frm"];

  /// <summary>The Windows code page vbadoc reads a file in when it isn't UTF-8, as the VBA editor writes them.</summary>
  public const int WindowsCodePage = 1252;

  const int Utf8CodePage = 65001;

  // Excel's document modules (the workbook, sheets and charts) have a base class id starting 000208 in VB_Base.
  const string DocumentBase = "0{000208";

  /// <exception cref="ExtractionException">The folder has no module files, or can't be read.</exception>
  public static ModuleFolder Read(string folder, bool recursive, bool folderGroups)
  {
    ArgumentException.ThrowIfNullOrWhiteSpace(folder);
    IReadOnlyList<SourceFolderFile> files;
    try
    {
      files = SourceFolder.Find(folder, [.. Extensions], recursive);
    }
    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
    {
      throw new ExtractionException($"Can't read the folder {folder}: {ex.Message}", ex);
    }

    if (files.Count == 0)
    {
      throw new ExtractionException(recursive
        ? $"No .bas, .cls or .frm files were found in {folder} or below it."
        : $"No .bas, .cls or .frm files were found in {folder}. Add --recursive to read its subfolders too.");
    }

    var warnings = new List<string>();
    var read = new List<(SourceFolderFile File, byte[] Bytes)>();
    foreach (var file in files)
    {
      try
      {
        read.Add((file, File.ReadAllBytes(file.FullPath)));
      }
      catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
      {
        warnings.Add($"Left out {file.RelativePath}, which can't be read: {ex.Message}");
      }
    }

    var codePage = CodePageOf(read.Select(r => r.Bytes).ToList(), read.Select(r => r.File.RelativePath).ToList(), warnings);
    var encoding = CodePages.Get(codePage);

    var modules = new List<VbaModule>();
    var byName = new Dictionary<string, SourceFolderFile>(StringComparer.OrdinalIgnoreCase);
    var clashing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    foreach (var (file, bytes) in read)
    {
      var text = encoding.GetString(bytes);
      if (text.StartsWith('﻿')) text = text[1..];
      var fileName = Path.GetFileNameWithoutExtension(file.RelativePath);
      var name = VbName().Match(text) is { Success: true } match ? match.Groups["name"].Value : fileName;
      if (!name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
      {
        warnings.Add($"{file.RelativePath} holds the module {name}, which it's documented as, as the VBA editor would import it");
      }

      if (byName.TryGetValue(name, out var existing))
      {
        if (string.Equals(FolderOf(existing.RelativePath), FolderOf(file.RelativePath), StringComparison.OrdinalIgnoreCase))
        {
          warnings.Add($"Left out {file.RelativePath}: {existing.RelativePath} beside it is the module {name} already");
          continue;
        }

        var pathName = string.Join('/', file.Folders.Append(name));
        if (byName.TryGetValue(pathName, out var samePath))
        {
          warnings.Add($"Left out {file.RelativePath}: {samePath.RelativePath} beside it is the module {name} already");
          continue;
        }

        warnings.Add($"The module {name} is in both {existing.RelativePath} and {file.RelativePath}. Both are shown, the second as "
          + $"{pathName}, and code that uses {name} is taken to mean {existing.RelativePath}. A workbook can't hold both, so import won't take either");
        clashing.Add(name);
        clashing.Add(pathName);
        name = pathName;
      }

      byName[name] = file;
      var source = WithoutHeader(text);
      modules.Add(new VbaModule
      {
        Name = name,
        Kind = KindOf(Path.GetExtension(file.RelativePath), text),
        StreamName = name,
        SourceBytes = encoding.GetBytes(source),
        Source = source,
        SourceFile = file.RelativePath,
        FileBytes = bytes,
        FolderFromPath = folderGroups && file.Folders.Count > 0 ? string.Join('.', file.Folders) : null,
      });
    }

    var project = new VbaProject
    {
      Name = Path.GetFileName(Path.TrimEndingDirectorySeparator(Path.GetFullPath(folder))),
      CodePage = codePage,
      Modules = [.. modules.OrderBy(m => m.Name, StringComparer.OrdinalIgnoreCase)],
      FromFolder = true,
      ClashingModules = clashing,
    };
    return new ModuleFolder(project, warnings);
  }

  /// <summary>
  /// A folder's project as <c>import</c> takes it (step S3): each module as a workbook holds it, without its file's path
  /// or bytes, so a bundle has the files as the VBA editor imports them, CRLF and with a class's header, and in Windows-1252,
  /// which a VBA project uses where a folder of UTF-8 files doesn't say. A character Windows-1252 hasn't becomes <c>?</c>.
  /// </summary>
  public static VbaProject ForImport(VbaProject project)
  {
    ArgumentNullException.ThrowIfNull(project);
    var windows = CodePages.Get(WindowsCodePage);
    return new VbaProject
    {
      Name = project.Name,
      CodePage = WindowsCodePage,
      Modules = [.. project.Modules.Select(m => new VbaModule
      {
        Name = m.Name,
        Kind = m.Kind,
        StreamName = m.StreamName,
        SourceBytes = windows.GetBytes(m.Source),
        Source = windows.GetString(windows.GetBytes(m.Source)),
      })],
      FromFolder = true,
      ClashingModules = project.ClashingModules,
    };
  }

  /// <summary>
  /// The code page to read every file in: UTF-8 when every file is valid UTF-8 and one has more than ASCII in
  /// it, otherwise Windows-1252, as the VBA editor writes. A folder mixing the two is read as Windows-1252, with a warning.
  /// </summary>
  static int CodePageOf(IReadOnlyList<byte[]> contents, IReadOnlyList<string> paths, List<string> warnings)
  {
    var strict = new UTF8Encoding(false, throwOnInvalidBytes: true);
    var utf8 = new List<string>();
    var other = 0;
    for (var i = 0; i < contents.Count; i++)
    {
      if (contents[i].All(b => b < 0x80)) continue;
      try
      {
        strict.GetString(contents[i]);
        utf8.Add(paths[i]);
      }
      catch (DecoderFallbackException)
      {
        other++;
      }
    }

    if (utf8.Count > 0 && other == 0) return Utf8CodePage;
    if (utf8.Count > 0)
    {
      warnings.Add($"Some files are UTF-8 and others aren't, so all are read as Windows-1252, and characters past ASCII in these may "
        + $"look wrong: {string.Join(", ", utf8)}");
    }

    return WindowsCodePage;
  }

  // A .cls file is a class unless its VB_Base says it's one of Excel's document modules.
  static VbaModuleKind KindOf(string extension, string text) => extension.ToLowerInvariant() switch
  {
    ".bas" => VbaModuleKind.Standard,
    ".frm" => VbaModuleKind.Form,
    _ => VbBase().Match(text) is { Success: true } vbBase && vbBase.Groups["base"].Value.StartsWith(DocumentBase, StringComparison.OrdinalIgnoreCase)
      ? VbaModuleKind.Document
      : VbaModuleKind.Class,
  };

  /// <summary>
  /// The module's source as a workbook stores it: without the VBA editor's <c>VERSION</c> header, or a form's designer
  /// block, which end where the <c>Attribute VB_Name</c> line starts, and with CRLF line endings.
  /// </summary>
  static string WithoutHeader(string text)
  {
    var crlf = text.ReplaceLineEndings("\r\n");
    if (!crlf.StartsWith("VERSION ", StringComparison.OrdinalIgnoreCase)) return crlf;
    var at = VbName().Match(crlf);
    return at.Success ? crlf[at.Index..] : crlf;
  }

  static string FolderOf(string relativePath) =>
    relativePath.LastIndexOf('/') is var at and >= 0 ? relativePath[..at] : string.Empty;

  [GeneratedRegex(@"^Attribute\s+VB_Name\s*=\s*""(?<name>[^""]+)""", RegexOptions.Multiline | RegexOptions.IgnoreCase)]
  private static partial Regex VbName();

  [GeneratedRegex(@"^Attribute\s+VB_Base\s*=\s*""(?<base>[^""]+)""", RegexOptions.Multiline | RegexOptions.IgnoreCase)]
  private static partial Regex VbBase();
}
