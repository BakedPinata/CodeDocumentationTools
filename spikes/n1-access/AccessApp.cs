using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace AccessSpike;

/// <summary>
/// Access itself, started for this run alone and driven through its automation interface: hidden, not under the
/// user's control, and closed again whatever happens. It's found by its registered name, with nothing to install or
/// reference, which is how accessdoc would have to do it.
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class AccessApp : IDisposable
{
  // Access's own numbers (the Ac… and dao constants), which a late-bound caller has to spell out.
  public const int AcForm = 2, AcMacro = 4, AcModule = 5, AcReport = 3, AcQuery = 1, AcTable = 0;
  public const int AcImport = 0, AcQuitSaveNone = 2, AcSaveYes = 1, AcDesign = 1, AcHidden = 1;
  public const int MsoAutomationSecurityForceDisable = 3;

  private readonly HashSet<int> before;

  public AccessApp()
  {
    before = [.. Process.GetProcessesByName("MSACCESS").Select(p => p.Id)];
    var type = Type.GetTypeFromProgID("Access.Application")
      ?? throw new InvalidOperationException("Access isn't installed here: nothing is registered as Access.Application.");
    App = Activator.CreateInstance(type) ?? throw new InvalidOperationException("Access didn't start.");
    App.Visible = false;
    App.UserControl = false;
  }

  /// <summary>
  /// Starts Access, does the work, and closes it. The work is given Access.Application, and everything it holds of
  /// Access has to be let go when it returns: Access stays running for as long as anything outside it holds one of
  /// its objects, so the work is kept in a call of its own, whose variables are dead by the time Access is asked to
  /// quit.
  /// </summary>
  public static void Use(Action<object> work)
  {
    var access = new AccessApp();
    try
    {
      Work(access, work);
    }
    finally
    {
      access.Dispose();
    }
  }

  [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
  private static void Work(AccessApp access, Action<object> work) => work(access.App);

  /// <summary>The Access.Application object.</summary>
  public dynamic App { get; }

  /// <summary>Access's version, such as 16.0, and whether it's the 64-bit one.</summary>
  public string Describe()
  {
    try
    {
      return $"Access {App.Version} (build {App.Build})";
    }
    catch (Exception)
    {
      return "Access (version not given)";
    }
  }

  /// <summary>Closes Access without saving, and makes sure the process this started has gone.</summary>
  public void Dispose()
  {
    // Let go of whatever the work held, so nothing outside Access keeps it running.
    GC.Collect();
    GC.WaitForPendingFinalizers();
    try
    {
      App.Quit(AcQuitSaveNone);
    }
    catch (Exception exception)
    {
      // Already gone, or refusing: the process is dealt with below.
      Console.WriteLine($"  Asking Access to quit failed: {exception.GetBaseException().Message}");
    }

    try
    {
      Marshal.FinalReleaseComObject(App);
    }
    catch (Exception)
    {
      // Nothing left to release.
    }

    GC.Collect();
    GC.WaitForPendingFinalizers();

    // Only a process that wasn't there before this started one is this run's to end.
    var deadline = Stopwatch.StartNew();
    while (deadline.ElapsedMilliseconds < 10_000)
    {
      var mine = Process.GetProcessesByName("MSACCESS").Where(p => !before.Contains(p.Id)).ToList();
      if (mine.Count == 0)
      {
        return;
      }

      Thread.Sleep(250);
    }

    foreach (var process in Process.GetProcessesByName("MSACCESS").Where(p => !before.Contains(p.Id)))
    {
      Console.WriteLine($"  Access (process {process.Id}) didn't close when asked, so it was ended.");
      try
      {
        process.Kill();
      }
      catch (Exception)
      {
        // It closed in the meantime.
      }
    }
  }
}
