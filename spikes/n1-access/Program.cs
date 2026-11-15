using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using AccessSpike;

// Step N1 of the generalisation plan: throwaway. See README.md.
//
//   accessspike build-sample <new.accdb>          build the sample database through Access
//   accessspike grow <database> <new.accdb> <n>   a copy with its forms copied until there are n
//   accessspike try <database> [options]          try both ways of reading it, and report
//     --keep <folder>      keep the text Access wrote, in this folder (it's deleted otherwise)
//     --names              list the objects' names in the report (leave it off for a real database)
//     --as-a-person-would  also open it the ordinary way, letting its start-up run: only for the sample

[assembly: SupportedOSPlatform("windows")]

if (args.Length == 0)
{
  Console.WriteLine("accessspike build-sample <new.accdb> | grow <database> <new.accdb> <forms> | try <database> [--keep <folder>] [--names] [--as-a-person-would]");
  return 2;
}

try
{
  switch (args[0])
  {
    case "build-sample" when args.Length == 2:
      SampleBuilder.Build(Path.Combine(AppContext.BaseDirectory, "sample"), args[1]);
      Console.WriteLine($"Built {Path.GetFullPath(args[1])}, with its linked table's back end at {SampleBuilder.BackEndPath}.");
      return 0;
    case "grow" when args.Length == 4:
      Grow(Path.GetFullPath(args[1]), Path.GetFullPath(args[2]), int.Parse(args[3]));
      return 0;
    case "try" when args.Length >= 2:
      Try(Path.GetFullPath(args[1]), args);
      return 0;
    default:
      Console.WriteLine("Unknown command, or the wrong number of arguments.");
      return 2;
  }
}
catch (Exception exception)
{
  Console.WriteLine($"Failed: {exception.GetBaseException().Message}");
  return 1;
}

// A copy of a database with one of its forms copied until it has n: the size of a real one, to time.
static void Grow(string database, string output, int count)
{
  var work = Path.Combine(Path.GetTempPath(), $"accessspike-{Guid.NewGuid():N}.accdb");
  File.Delete(output);
  AccessApp.Use(application =>
  {
    dynamic app = application;
    // Into a new database, by import, so the sample's start-up doesn't run and isn't carried over as start-up.
    app.NewCurrentDatabase(output);
    var source = app.DBEngine.OpenDatabase(database, false, true);
    var forms = new List<string>();
    var others = new List<(int Type, string Name)>();
    foreach (var (container, type) in new[] { ("Forms", AccessApp.AcForm), ("Modules", AccessApp.AcModule), ("Scripts", AccessApp.AcMacro) })
    {
      var documents = source.Containers[container].Documents;
      for (var i = 0; i < (int)documents.Count; i++)
      {
        string name = documents[i].Name;
        if (type == AccessApp.AcForm) forms.Add(name); else if (name != "AutoExec") others.Add((type, name));
      }
    }

    var tables = new List<string>();
    for (var i = 0; i < (int)source.TableDefs.Count; i++)
    {
      string name = source.TableDefs[i].Name;
      if (!name.StartsWith("MSys", StringComparison.OrdinalIgnoreCase)) tables.Add(name);
    }

    source.Close();
    foreach (var name in tables) app.DoCmd.TransferDatabase(AccessApp.AcImport, "Microsoft Access", database, AccessApp.AcTable, name, name, false);
    foreach (var (type, name) in others) app.DoCmd.TransferDatabase(AccessApp.AcImport, "Microsoft Access", database, type, name, name, false);
    foreach (var name in forms) app.DoCmd.TransferDatabase(AccessApp.AcImport, "Microsoft Access", database, AccessApp.AcForm, name, name, false);
    for (var n = forms.Count; n < count; n++)
    {
      var original = forms[n % forms.Count];
      app.DoCmd.CopyObject(Type.Missing, $"{original}_{n:000}", AccessApp.AcForm, original);
    }

    app.CloseCurrentDatabase();
  });

  File.Delete(work);
  Console.WriteLine($"Built {output} with {count} forms.");
}

static void Try(string database, string[] args)
{
  var names = args.Contains("--names");
  var keepAt = Array.IndexOf(args, "--keep");
  var keep = keepAt >= 0 && keepAt + 1 < args.Length ? Path.GetFullPath(args[keepAt + 1]) : null;
  var output = keep ?? Path.Combine(Path.GetTempPath(), $"accessspike-text-{Guid.NewGuid():N}");

  Console.WriteLine($"accessspike, step N1. {RuntimeInformation.OSDescription}, {RuntimeInformation.ProcessArchitecture} process.");
  Console.WriteLine($"Database: {(names ? database : "(name withheld)")}, {new FileInfo(database).Length / 1024:N0} KB");
  Console.WriteLine();

  var methods = new List<ExportMethod> { ExportMethod.Import, ExportMethod.OpenWithCodeDisabled };
  if (args.Contains("--as-a-person-would")) methods.Add(ExportMethod.OpenAsAPersonWould);
  ExportResult? best = null;
  foreach (var method in methods)
  {
    Console.WriteLine($"Choice A, Access writes text from a copy: {Describe(method)}");
    var folder = Path.Combine(output, method.ToString());
    if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
    var result = Exporter.Export(database, method, folder);
    if (result.Failure is not null) Console.WriteLine($"  FAILED: {result.Failure}");
    Console.WriteLine($"  Got: {result.Forms} forms, {result.Reports} reports, {result.Modules} modules, {result.Macros} macros, {result.Queries} queries, " +
      $"{result.LocalTables} local and {result.LinkedTables} linked tables; {result.TextBytes / 1024:N0} KB of text; {result.Code.Count} modules' code (form and report modules included)");
    Console.WriteLine($"  Time: starting Access {Seconds(result.Starting)}, listing {Seconds(result.Reading)}, writing text {Seconds(result.Writing)}, closing {Seconds(result.Closing)}");
    Console.WriteLine($"  Start-up: {result.StartUp}");
    Console.WriteLine($"  Access processes left running by this: {AccessProcessesNow()}");
    if (result.Failure is null && method != ExportMethod.OpenAsAPersonWould) best ??= result;
    Console.WriteLine();
  }

  Console.WriteLine("Choice C, the file read directly, with no Access:");
  try
  {
    var clock = Stopwatch.StartNew();
    var direct = DirectReader.Read(database);
    Console.WriteLine($"  Got: {direct.Modules.Count} modules' code (form and report modules included), in {Seconds(clock.Elapsed)}");
    if (best is not null)
    {
      var fromAccess = best.Code.ToDictionary(m => m.Name, m => Code(m.Text), StringComparer.OrdinalIgnoreCase);
      var same = direct.Modules.Count(m => fromAccess.TryGetValue(m.Name, out var text) && text == Code(m.Text));
      var missing = fromAccess.Keys.Except(direct.Modules.Select(m => m.Name), StringComparer.OrdinalIgnoreCase).Count();
      Console.WriteLine($"  Against choice A's code: {same} the same, {direct.Modules.Count - same} different or not in A, {missing} in A and not found directly");
      if (names)
      {
        foreach (var module in direct.Modules.Where(m => !fromAccess.TryGetValue(m.Name, out var text) || text != Code(m.Text)))
        {
          Console.WriteLine($"    differs: {module.Name}");
          // The first line that differs, to say what kind of difference it is.
          var theirs = fromAccess.TryGetValue(module.Name, out var a) ? a.Split((char)10) : [];
          var ours = Code(module.Text).Split((char)10);
          var at = Enumerable.Range(0, Math.Max(theirs.Length, ours.Length)).FirstOrDefault(i => i >= theirs.Length || i >= ours.Length || theirs[i] != ours[i]);
          Console.WriteLine($"      line {at + 1}, from Access: {(at < theirs.Length ? theirs[at] : "(none)")}");
          Console.WriteLine($"      line {at + 1}, directly:    {(at < ours.Length ? ours[at] : "(none)")}");
        }
      }
    }

    var streams = direct.Storage.Where(s => s.Size >= 0).ToList();
    Console.WriteLine($"  The file's own store (MSysAccessStorage): {direct.Storage.Count} entries, {streams.Count} with data, {streams.Sum(s => (long)s.Size) / 1024:N0} KB");
    foreach (var group in direct.Storage.GroupBy(s => string.Join("/", s.Path.Split('/').Take(2))).OrderBy(g => g.Key, StringComparer.Ordinal))
    {
      Console.WriteLine($"    {group.Key}: {group.Count()} entries, {group.Where(s => s.Size > 0).Sum(s => (long)s.Size) / 1024:N0} KB");
    }

    if (names)
    {
      // What a form is stored as: the names and sizes of its streams, which is all that can be said of them.
      foreach (var entry in direct.Storage.Where(s => s.Path.Contains("/Forms/", StringComparison.Ordinal)).Take(12))
      {
        Console.WriteLine($"    {entry.Path}: {(entry.Size < 0 ? "a storage" : entry.Size + " bytes")}");
      }
    }
  }
  catch (Exception exception)
  {
    Console.WriteLine($"  FAILED: {exception.GetBaseException().Message}");
  }

  if (keep is null && Directory.Exists(output)) Directory.Delete(output, recursive: true);
  else if (keep is not null) Console.WriteLine($"\nThe text is in {keep}.");
}

static string Describe(ExportMethod method) => method switch
{
  ExportMethod.Import => "a new, empty database imports each object from the copy (the database itself is never opened)",
  ExportMethod.OpenWithCodeDisabled => "the copy is opened with all code disabled (AutomationSecurity)",
  _ => "the copy is opened as a person would open it, so its start-up runs (for comparison)",
};

static string Seconds(TimeSpan time) => $"{time.TotalSeconds:0.0} s";

static int AccessProcessesNow() => Process.GetProcessesByName("MSACCESS").Length;

// Code as it compares: without the attribute lines Access writes in some places and not others, and without the
// differences in line endings and trailing space.
static string Code(string text) => string.Join("\n", text.Replace("\r\n", "\n").Split('\n')
  .Where(line => !line.StartsWith("Attribute ", StringComparison.Ordinal))
  .Select(line => line.TrimEnd())).Trim();
