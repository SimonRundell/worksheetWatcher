using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace WorksheetWatcher;

/// <summary>
/// Small per-user preferences stored at <c>%AppData%\WorksheetWatcher\ui.json</c>,
/// separate from <c>WorksheetWatcher.config.json</c> (which sits next to the executable
/// and is not writable when the app is installed under Program Files).
/// </summary>
public sealed class UserSettings
{
    private const string FileName = "ui.json";

    /// <summary>
    /// Interface scale as a percentage. 100 is the design size; the View menu offers a
    /// few steps either side. Clamped to a sensible range on load.
    /// </summary>
    [JsonPropertyName("interfaceScalePercent")]
    public int InterfaceScalePercent { get; set; } = 100;

    /// <summary>The last notebook picked, so it is pre-selected next time it is open.</summary>
    [JsonPropertyName("lastNotebookId")]
    public string? LastNotebookId { get; set; }

    /// <summary>The last worksheet title watched, pre-selected once its notebook loads.</summary>
    [JsonPropertyName("lastWorksheetTitle")]
    public string? LastWorksheetTitle { get; set; }

    [JsonIgnore]
    private static string Path =>
        System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "WorksheetWatcher", FileName);

    /// <summary>The allowed scale steps offered in the UI.</summary>
    public static readonly int[] ScaleSteps = { 90, 100, 110, 125, 150, 175, 200 };

    /// <summary>Loads settings, returning defaults if the file is missing or unreadable.</summary>
    public static UserSettings Load()
    {
        try
        {
            if (File.Exists(Path))
            {
                var s = JsonSerializer.Deserialize<UserSettings>(File.ReadAllText(Path));
                if (s is not null)
                {
                    s.InterfaceScalePercent = Clamp(s.InterfaceScalePercent);
                    return s;
                }
            }
        }
        catch
        {
            // fall through to defaults
        }
        return new UserSettings();
    }

    /// <summary>Writes settings, swallowing I/O errors (a failed save is not worth a dialog).</summary>
    public void Save()
    {
        try
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(Path)!);
            File.WriteAllText(Path, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch
        {
            // ignored
        }
    }

    /// <summary>Snaps an arbitrary percentage to the nearest allowed step.</summary>
    public static int Clamp(int percent)
    {
        if (percent <= ScaleSteps[0]) return ScaleSteps[0];
        if (percent >= ScaleSteps[^1]) return ScaleSteps[^1];
        return ScaleSteps.OrderBy(s => Math.Abs(s - percent)).First();
    }
}
