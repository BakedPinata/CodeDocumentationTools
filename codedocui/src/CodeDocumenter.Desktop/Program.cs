using CodeDocumenter.Desktop.Errors;
using CodeDocumenter.Desktop.Settings;
using CodeDocumenter.Desktop.Window;

namespace CodeDocumenter.Desktop;

internal static class Program
{
  /// <summary>
  /// Opens the window. A workbook or folder given on the command line is chosen to start with. An error the app doesn't
  /// expect, anywhere, is reported once and closes the app (<see cref="ErrorHandler"/>).
  /// </summary>
  [STAThread]
  private static void Main(string[] args)
  {
    ApplicationConfiguration.Initialize();
    MainWindow? window = null;
    var errors = new ErrorHandler(
      ReportFolder.DefaultPath,
      MainWindow.Versions,
      () => window is null ? ["The window wasn't open yet."]
        : window.InvokeRequired ? ["Not read: the error came on another thread, and only the window's own can read it."]
        : window.DescribeState(),
      new ErrorDialog(() => window),
      () => SettingsStore.StartAfresh(SettingsStore.DefaultPath),
      Environment.Exit);

    Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
    Application.ThreadException += (_, e) => errors.Handle(e.Exception, onWindowThread: true);
    AppDomain.CurrentDomain.UnhandledException += (_, e) => OnAnotherThread(errors, e.ExceptionObject);
    try
    {
      window = new MainWindow(SettingsStore.DefaultPath, args.Length > 0 ? args[0] : null);
      Application.Run(window);
    }
    catch (Exception ex)
    {
      // Making the window failed, before there was one to catch it.
      errors.Handle(ex, onWindowThread: true);
    }
    finally
    {
      window?.Dispose();
    }
  }

  // An error on another thread ends the app as soon as this returns, so it's reported first, from a thread of the
  // window's kind (STA), which the dialog's Copy needs, while this one waits. If it comes while another is being
  // reported, it's counted and this thread waits for the app to close after that one.
  private static void OnAnotherThread(ErrorHandler errors, object error)
  {
    var exception = error as Exception ?? new InvalidOperationException("Something that isn't an exception was thrown.");
    var reporter = new Thread(() => errors.Handle(exception, onWindowThread: false));
    reporter.SetApartmentState(ApartmentState.STA);
    reporter.Start();
    reporter.Join();
    Thread.Sleep(Timeout.Infinite);
  }
}
