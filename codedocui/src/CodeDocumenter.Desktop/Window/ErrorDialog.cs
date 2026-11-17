using System.ComponentModel;
using System.Diagnostics;
using CodeDocumenter.Desktop.Errors;

namespace CodeDocumenter.Desktop.Window;

/// <summary>
/// An error in a task dialog: what it was, with Open the report, Copy the report, Close codedocui, and a tick box to
/// start afresh next time. Over the window when it's on the window's thread, or on its own otherwise.
/// </summary>
/// <param name="window">The window, or null before it's made.</param>
internal sealed class ErrorDialog(Func<Control?> window) : IErrorDialog
{
  public bool Show(ShownError error)
  {
    ArgumentNullException.ThrowIfNull(error);

    // These two leave the dialog open.
    var open = new TaskDialogButton("&Open the report") { AllowCloseDialog = false, Visible = error.ReportPath is not null };
    open.Click += (_, _) => OpenReport(error.ReportPath!);
    var copy = new TaskDialogButton("Co&py the report") { AllowCloseDialog = false };
    copy.Click += (_, _) => Clipboard.SetText(error.Report);
    var close = new TaskDialogButton("&Close codedocui");
    var afresh = new TaskDialogVerificationCheckBox("&Start afresh next time: forget how the tabs were left, and the recent workbooks and folders");

    var saved = error.ReportPath is null
      ? "Its report couldn't be saved, but Copy the report copies it."
      : "Its report is saved, to read or pass on.";
    var page = new TaskDialogPage
    {
      Caption = "codedocui",
      Heading = "Something went wrong, and codedocui has to close",
      Text = $"{error.Summary}.\n\n{saved} It names no file or folder and says nothing of what's in them.",
      Footnote = new TaskDialogFootnote($"Reports are kept in %LOCALAPPDATA%\\codedocui\\reports, the newest {ReportFolder.Kept}."),
      Icon = TaskDialogIcon.Error,
      AllowCancel = true,
      Verification = afresh,
      DefaultButton = close,
    };
    page.Buttons.Add(open);
    page.Buttons.Add(copy);
    page.Buttons.Add(close);

    // A window can only own a dialog on its own thread.
    if (window() is { IsHandleCreated: true, IsDisposed: false, InvokeRequired: false } owner)
    {
      TaskDialog.ShowDialog(owner, page);
    }
    else
    {
      TaskDialog.ShowDialog(page, TaskDialogStartupLocation.CenterScreen);
    }

    return afresh.Checked;
  }

  private static void OpenReport(string path)
  {
    try
    {
      Process.Start(new ProcessStartInfo(path) { UseShellExecute = true })?.Dispose();
    }
    catch (Exception ex) when (ex is Win32Exception or InvalidOperationException or FileNotFoundException)
    {
      // Copy the report is still there.
    }
  }
}
