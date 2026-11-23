using System.Text;
using CodeDocumenter.Shared;
using VbaDocumenter.Core.Extraction;
using VbaDocumenter.Core.Output;

namespace VbaDocumenter.Tests.Extraction;

/// <summary>Reading a folder of exported modules, as the VBA editor, Rubberduck or someone else left it.</summary>
public sealed class ModuleFolderReaderTests : IDisposable
{
  static ModuleFolderReaderTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

  readonly string _folder = Path.Combine(Path.GetTempPath(), $"vbadoc-folder-{Guid.NewGuid():N}");

  public ModuleFolderReaderTests() => Directory.CreateDirectory(_folder);

  public void Dispose()
  {
    if (Directory.Exists(_folder)) Directory.Delete(_folder, recursive: true);
  }

  void Write(string path, string text, Encoding? encoding = null)
  {
    var full = Path.Combine(_folder, path);
    Directory.CreateDirectory(Path.GetDirectoryName(full)!);
    File.WriteAllBytes(full, (encoding ?? Encoding.ASCII).GetBytes(text));
  }

  const string EditorClass =
    "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Logger\"\r\nAttribute VB_GlobalNameSpace = False\r\n"
    + "Attribute VB_Base = \"0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}\"\r\nOption Explicit\r\n\r\nPublic Sub Log(ByVal text As String)\r\nEnd Sub\r\n";

  const string EditorSheet =
    "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Sheet1\"\r\n"
    + "Attribute VB_Base = \"0{00020820-0000-0000-C000-000000000046}\"\r\nOption Explicit\r\n";

  [Fact]
  public void Modules_are_named_by_VB_Name_and_kinded_by_extension_and_VB_Base()
  {
    Write("Main.bas", "Attribute VB_Name = \"Main\"\r\nPublic Sub Run()\r\n  Logger.Log \"x\"\r\nEnd Sub\r\n");
    Write("Logger.cls", EditorClass);
    Write("Sheet1.cls", EditorSheet);
    Write("Renamed.bas", "Attribute VB_Name = \"Helpers\"\r\n");
    Write("notes.txt", "not a module");

    var read = ModuleFolderReader.Read(_folder, recursive: false, folderGroups: false);

    Assert.Equal(["Helpers", "Logger", "Main", "Sheet1"], read.Project.Modules.Select(m => m.Name));
    Assert.Equal([VbaModuleKind.Standard, VbaModuleKind.Class, VbaModuleKind.Standard, VbaModuleKind.Document], read.Project.Modules.Select(m => m.Kind));
    Assert.Equal("Renamed.bas", read.Project.Modules[0].SourceFile);
    Assert.Contains(read.Warnings, w => w.Contains("Renamed.bas holds the module Helpers", StringComparison.Ordinal));
    Assert.True(read.Project.FromFolder);
    Assert.Empty(read.Project.References);

    // The page and the parser see the file as it is; import sees the source a workbook holds, without the header.
    var logger = read.Project.Modules[1];
    Assert.Equal(File.ReadAllBytes(Path.Combine(_folder, "Logger.cls")), ProjectExporter.FileContent(logger, 1252));
    Assert.StartsWith("Attribute VB_Name", logger.Source, StringComparison.Ordinal);
  }

  [Fact]
  public void Files_from_elsewhere_are_read_whatever_their_line_endings_header_or_code_page()
  {
    Write("Accents.bas", "Attribute VB_Name = \"Accents\"\nPublic Const Town = \"Zürich\"\n", Encoding.GetEncoding(1252));
    Write("Bare.cls", "Option Explicit\nPublic Name As String\n");
    Write("NoName.bas", "Public Sub Hello()\nEnd Sub\n");

    var read = ModuleFolderReader.Read(_folder, recursive: false, folderGroups: false);

    Assert.Equal(1252, read.Project.CodePage);
    Assert.Contains("Zürich", read.Project.Modules.Single(m => m.Name == "Accents").Source, StringComparison.Ordinal);
    Assert.Equal(VbaModuleKind.Class, read.Project.Modules.Single(m => m.Name == "Bare").Kind);
    Assert.Equal("NoName", read.Project.Modules.Single(m => m.SourceFile == "NoName.bas").Name);
    Assert.Contains("\r\n", read.Project.Modules[0].Source, StringComparison.Ordinal);
  }

  [Fact]
  public void A_folder_of_UTF8_files_is_read_as_UTF8_with_or_without_a_byte_order_mark()
  {
    Write("Plain.bas", "Attribute VB_Name = \"Plain\"\r\nPublic Const Town = \"Köln\"\r\n", new UTF8Encoding(false));
    Write("Marked.bas", "Attribute VB_Name = \"Marked\"\r\nPublic Const Town = \"Zürich\"\r\n", new UTF8Encoding(true));
    File.WriteAllBytes(Path.Combine(_folder, "Marked.bas"), [.. Encoding.UTF8.GetPreamble(), .. Encoding.UTF8.GetBytes("Attribute VB_Name = \"Marked\"\r\nPublic Const Town = \"Zürich\"\r\n")]);

    var read = ModuleFolderReader.Read(_folder, recursive: false, folderGroups: false);

    Assert.Equal(65001, read.Project.CodePage);
    Assert.Equal(["Marked", "Plain"], read.Project.Modules.Select(m => m.Name));
    Assert.Contains("Zürich", read.Project.Modules[0].Source, StringComparison.Ordinal);
  }

  [Fact]
  public void Subfolders_are_read_with_recursive_and_become_folders_only_with_folder_groups()
  {
    Write("Main.bas", "Attribute VB_Name = \"Main\"\r\n");
    Write("Util/Text/Strings.bas", "Attribute VB_Name = \"Strings\"\r\n");
    Write(".git/Hidden.bas", "Attribute VB_Name = \"Hidden\"\r\n");

    Assert.Single(ModuleFolderReader.Read(_folder, recursive: false, folderGroups: false).Project.Modules);
    var plain = ModuleFolderReader.Read(_folder, recursive: true, folderGroups: false).Project.Modules;
    var grouped = ModuleFolderReader.Read(_folder, recursive: true, folderGroups: true).Project.Modules;

    Assert.Equal(["Main", "Strings"], plain.Select(m => m.Name));
    Assert.Equal("Util/Text/Strings.bas", plain[1].SourceFile);
    Assert.Null(plain[1].FolderFromPath);
    Assert.Equal("Util.Text", grouped[1].FolderFromPath);
    Assert.Null(grouped[0].FolderFromPath);
  }

  [Fact]
  public void The_same_name_in_two_folders_keeps_both_and_marks_them_as_clashing()
  {
    Write("New/Utils.bas", "Attribute VB_Name = \"Utils\"\r\n");
    Write("Old/Utils.bas", "Attribute VB_Name = \"Utils\"\r\n");
    Write("Old/Same.bas", "Attribute VB_Name = \"Utils\"\r\n");

    var read = ModuleFolderReader.Read(_folder, recursive: true, folderGroups: false);

    Assert.Equal(["Old/Utils", "Utils"], read.Project.Modules.Select(m => m.Name));
    Assert.Equal(["Old/Utils", "Utils"], read.Project.ClashingModules.Order(StringComparer.Ordinal));
    Assert.Contains(read.Warnings, w => w.Contains("import won't take either", StringComparison.Ordinal));
  }

  [Fact]
  public void A_folder_with_no_module_files_says_so()
  {
    Write("Deep/Main.bas", "Attribute VB_Name = \"Main\"\r\n");

    var thrown = Assert.Throws<ExtractionException>(() => ModuleFolderReader.Read(_folder, recursive: false, folderGroups: false));
    Assert.Contains("--recursive", thrown.Message, StringComparison.Ordinal);
  }
}
