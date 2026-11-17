using System.Text.Json;
using System.Text.Json.Serialization;
using CodeDocumenter.Shared.Files;

namespace CodeDocumenter.Desktop.Settings;

/// <summary>Reads and writes the app's settings, a JSON file for each person using the machine.</summary>
internal static class SettingsStore
{
  private static readonly JsonSerializerOptions Options = new()
  {
    WriteIndented = true,
    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
  };

  /// <summary>
  /// <c>%LOCALAPPDATA%\codedocui\settings.json</c>: per person and per machine, where Windows Forms keeps its own settings.
  /// </summary>
  public static string DefaultPath { get; } =
    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "codedocui", "settings.json");

  /// <summary>The settings, or the defaults if there are none yet or they can't be read.</summary>
  public static AppSettings Load(string path)
  {
    ArgumentNullException.ThrowIfNull(path);
    try
    {
      return File.Exists(path)
        ? (JsonSerializer.Deserialize<AppSettings>(File.ReadAllBytes(path), Options) ?? new AppSettings()).Normalise()
        : new AppSettings();
    }
    catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException or NotSupportedException)
    {
      // Settings are a convenience: ones that can't be read are started again rather than stopping the app.
      return new AppSettings();
    }
  }

  /// <summary>Writes the settings, all at once, so a failure leaves the previous ones.</summary>
  public static void Save(string path, AppSettings settings)
  {
    ArgumentNullException.ThrowIfNull(path);
    ArgumentNullException.ThrowIfNull(settings);
    Directory.CreateDirectory(Path.GetDirectoryName(path)!);
    AtomicFileWriter.Write(path, JsonSerializer.SerializeToUtf8Bytes(settings, Options));
  }

  /// <summary>
  /// Puts the settings aside as <c>settings.old.json</c>, in place of any put aside before, so the app starts from the
  /// defaults next time and they can still be looked at.
  /// </summary>
  public static void StartAfresh(string path)
  {
    ArgumentNullException.ThrowIfNull(path);
    if (File.Exists(path))
    {
      File.Move(path, OldPath(path), overwrite: true);
    }
  }

  /// <summary>Where <see cref="StartAfresh"/> puts the settings.</summary>
  public static string OldPath(string path) => Path.ChangeExtension(path, ".old.json");
}
