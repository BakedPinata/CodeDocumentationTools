using VbaDocumenter.Core.Extraction;
using VbaDocumenter.Core.Output;

namespace VbaDocumenter.Core.Parsing;

public static class ProjectParser
{
  /// <summary>
  /// Parses every module. Each module is parsed from the exact text of its exported file, so member line
  /// numbers match the files in the output folder (and the source shown in the page).
  /// </summary>
  public static ProjectModel Parse(VbaProject project)
  {
    var encoding = CodePages.Get(project.CodePage);
    return new ProjectModel
    {
      Name = project.Name,
      CodePage = project.CodePage,
      Modules = project.Modules
            .Select(m => ModuleParser.Parse(m.Name, m.Kind, FileText(encoding, m), m.FolderFromPath))
            .ToList(),
    };
  }

  /// <summary>
  /// A module's file as text. A file from a folder may start with a UTF-8 byte order mark, which isn't part of a line,
  /// so it's left out, as a browser's decoder leaves it out.
  /// </summary>
  public static string FileText(System.Text.Encoding encoding, VbaModule module)
  {
    var text = encoding.GetString(ProjectExporter.FileContent(module));
    return text.StartsWith('﻿') ? text[1..] : text;
  }
}
