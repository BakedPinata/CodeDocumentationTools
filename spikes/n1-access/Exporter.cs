using System.Diagnostics;
using System.Runtime.Versioning;

namespace AccessSpike;

/// <summary>How choice A gets at the database's objects.</summary>
public enum ExportMethod
{
  /// <summary>
  /// Never opens the database: a new, empty database imports each object from the copy, and writes it as text from
  /// there. Nothing of the database's start-up can run, as the database is never the one that's open.
  /// </summary>
  Import,

  /// <summary>Opens the copy with Access told to disable all code first (AutomationSecurity).</summary>
  OpenWithCodeDisabled,

  /// <summary>Opens the copy as a person would. Start-up runs. Only for the sample, to show the mark it leaves.</summary>
  OpenAsAPersonWould,
}

/// <summary>What one run of choice A got.</summary>
public sealed record ExportResult(
  ExportMethod Method, string? Failure, int Forms, int Reports, int Modules, int Macros, int Queries, int LocalTables, int LinkedTables,
  TimeSpan Starting, TimeSpan Reading, TimeSpan Writing, TimeSpan Closing, string StartUp, IReadOnlyList<ModuleText> Code, long TextBytes);

/// <summary>
/// Choice A: Access writes each form, report, module and macro as text (its own SaveAsText) from a copy of the
/// database, hidden, and is closed afterwards. The database given is only ever copied.
/// </summary>
[SupportedOSPlatform("windows")]
public static class Exporter
{
  private const int DbSystemObject = unchecked((int)0x80000002);

  public static ExportResult Export(string databasePath, ExportMethod method, string outputFolder)
  {
    var work = Path.Combine(Path.GetTempPath(), $"accessspike-{Guid.NewGuid():N}");
    Directory.CreateDirectory(work);
    Directory.CreateDirectory(outputFolder);
    var copy = Path.Combine(work, "copy.accdb");
    File.Copy(databasePath, copy);
    File.SetAttributes(copy, FileAttributes.Normal);

    var forms = new List<string>();
    var reports = new List<string>();
    var modules = new List<string>();
    var macros = new List<string>();
    var queries = 0;
    var localTables = 0;
    var linkedTables = 0;
    var startUp = "";
    string? failure = null;
    var clock = Stopwatch.StartNew();
    TimeSpan starting = default, reading = default, writing = default;

    try
    {
      AccessApp.Use(application =>
      {
        dynamic app = application;
        try
        {
          if (method == ExportMethod.Import)
          {
            app.NewCurrentDatabase(Path.Combine(work, "blank.accdb"));
          }
          else
          {
            if (method == ExportMethod.OpenWithCodeDisabled)
            {
              app.AutomationSecurity = AccessApp.MsoAutomationSecurityForceDisable;
            }

            app.OpenCurrentDatabase(copy, false);
          }

          starting = clock.Elapsed;

          // What's in the database, read through the database engine, which runs nothing.
          var source = method == ExportMethod.Import ? app.DBEngine.OpenDatabase(copy, false, true) : app.CurrentDb();
          var marksBefore = method == ExportMethod.Import ? Marks(source) : -1;
          forms = Names(source, "Forms");
          reports = Names(source, "Reports");
          modules = Names(source, "Modules");
          macros = Names(source, "Scripts");
          var sql = new List<(string Name, string Sql)>();
          for (var i = 0; i < (int)source.QueryDefs.Count; i++)
          {
            var query = source.QueryDefs[i];
            string name = query.Name;
            if (!name.StartsWith('~'))
            {
              sql.Add((name, (string)query.SQL));
            }
          }

          queries = sql.Count;
          var tables = new List<string>();
          for (var i = 0; i < (int)source.TableDefs.Count; i++)
          {
            var table = source.TableDefs[i];
            string name = table.Name;
            if (((int)table.Attributes & DbSystemObject) != 0 || name.StartsWith("MSys", StringComparison.OrdinalIgnoreCase) || name.StartsWith('~'))
            {
              continue;
            }

            string connect = table.Connect ?? "";
            if (connect.Length > 0) linkedTables++; else localTables++;
            tables.Add($"{name}\t{(connect.Length > 0 ? "linked" : "local")}\t{(int)table.Fields.Count} fields\t{WithoutSecrets(connect)}");
          }

          if (method == ExportMethod.Import)
          {
            source.Close();
          }

          reading = clock.Elapsed - starting;

          foreach (var (kind, type, names) in new[] { ("forms", AccessApp.AcForm, forms), ("reports", AccessApp.AcReport, reports), ("modules", AccessApp.AcModule, modules), ("macros", AccessApp.AcMacro, macros) })
          {
            var folder = Directory.CreateDirectory(Path.Combine(outputFolder, kind)).FullName;
            foreach (var name in names)
            {
              if (method == ExportMethod.Import)
              {
                app.DoCmd.TransferDatabase(AccessApp.AcImport, "Microsoft Access", copy, type, name, name, false);
              }

              app.SaveAsText(type, name, Path.Combine(folder, SafeFileName(name) + ".txt"));
            }
          }

          var queryFolder = Directory.CreateDirectory(Path.Combine(outputFolder, "queries")).FullName;
          foreach (var (name, text) in sql)
          {
            File.WriteAllText(Path.Combine(queryFolder, SafeFileName(name) + ".sql"), text);
          }

          File.WriteAllLines(Path.Combine(outputFolder, "tables.txt"), tables);
          writing = clock.Elapsed - starting - reading;

          // Did any of the database's own start-up run? The sample's leaves three marks.
          var open = (int)app.Forms.Count;
          var tempVars = (int)app.TempVars.Count;
          var marksAfter = -1;
          if (method == ExportMethod.Import)
          {
            var again = app.DBEngine.OpenDatabase(copy, false, true);
            marksAfter = Marks(again);
            again.Close();
          }
          else
          {
            marksAfter = Marks(app.CurrentDb());
          }

          startUp = $"{open} form(s) open, {tempVars} TempVars set" +
            (marksAfter < 0 ? "" : $", StartupLog rows {(marksBefore < 0 ? "" : marksBefore + " before, ")}{marksAfter} after");
          app.CloseCurrentDatabase();
        }
        catch (Exception exception)
        {
          failure = exception.GetBaseException().Message;
        }
      });
    }
    catch (Exception exception)
    {
      failure ??= exception.GetBaseException().Message;
    }

    var closing = clock.Elapsed - starting - reading - writing;
    var code = new List<ModuleText>();
    long textBytes = 0;
    foreach (var file in Directory.Exists(outputFolder) ? Directory.GetFiles(outputFolder, "*.txt", SearchOption.AllDirectories) : [])
    {
      textBytes += new FileInfo(file).Length;
      var kind = Path.GetFileName(Path.GetDirectoryName(file));
      var name = Path.GetFileNameWithoutExtension(file);
      var text = File.ReadAllText(file); // the byte order mark says which encoding Access wrote
      if (kind == "modules")
      {
        code.Add(new ModuleText(name, text));
      }
      else if (kind is "forms" or "reports")
      {
        var at = text.IndexOf("\nCodeBehindForm", StringComparison.Ordinal);
        if (at >= 0)
        {
          code.Add(new ModuleText((kind == "forms" ? "Form_" : "Report_") + name, text[(text.IndexOf('\n', at + 1) + 1)..]));
        }
      }
    }

    try
    {
      Directory.Delete(work, recursive: true);
    }
    catch (IOException)
    {
      Console.WriteLine($"  Couldn't delete {work}: Access may still hold the copy.");
    }

    return new ExportResult(method, failure, forms.Count, reports.Count, modules.Count, macros.Count, queries, localTables, linkedTables,
      starting, reading, writing, closing, startUp, code, textBytes);
  }

  private static List<string> Names(dynamic database, string container)
  {
    var names = new List<string>();
    var documents = database.Containers[container].Documents;
    for (var i = 0; i < (int)documents.Count; i++)
    {
      string name = documents[i].Name;
      if (!name.StartsWith('~'))
      {
        names.Add(name);
      }
    }

    return names;
  }

  /// <summary>The rows in the sample's StartupLog table, or -1 for a database without one.</summary>
  private static int Marks(dynamic database)
  {
    try
    {
      var rows = database.OpenRecordset("SELECT Count(*) AS N FROM StartupLog");
      int count = Convert.ToInt32(rows.Fields["N"].Value);
      rows.Close();
      return count;
    }
    catch (Exception)
    {
      return -1;
    }
  }

  /// <summary>A linked table's connection, without a user name or password.</summary>
  private static string WithoutSecrets(string connect) =>
    string.Join(";", connect.Split(';').Where(part =>
      !part.TrimStart().StartsWith("PWD=", StringComparison.OrdinalIgnoreCase) && !part.TrimStart().StartsWith("UID=", StringComparison.OrdinalIgnoreCase)));

  private static string SafeFileName(string name) =>
    string.Concat(name.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
}
