using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace WorksheetWatcher;

/// <summary>
/// Strongly typed view of <c>WorksheetWatcher.config.json</c>, which sits next to the
/// executable and is read once at start-up. Keeping locale/deployment sensitive values
/// (the Class Notebook system section-group names, the poll cadence) in config rather
/// than hard-coding them means the tool can be tuned per classroom without a recompile.
/// </summary>
public sealed class AppConfig
{
    /// <summary>
    /// Section groups sitting directly under the notebook whose names match one of
    /// these (case-insensitive, exact after normalisation) are treated as Class Notebook
    /// system areas and are never reported as students.
    /// </summary>
    [JsonPropertyName("excludedSectionGroupNames")]
    public List<string> ExcludedSectionGroupNames { get; set; } = new()
    {
        "Content Library",
        "Collaboration Space"
    };

    /// <summary>
    /// How often the background watcher re-checks each student's page for a new
    /// <c>lastModifiedTime</c> before deciding whether to re-render it. The UI only offers
    /// 5/10/15/30s presets; this is just the one pre-selected on start-up. Floored at 5
    /// seconds at run time regardless of what is configured - measured against a real
    /// 23-student notebook, the hierarchy re-check itself costs 50-200ms, so 5s is safe;
    /// the real pacing limit is how many students' pages actually changed that tick, since
    /// each one costs a separate ~1-2s Publish/rasterise call.
    /// </summary>
    [JsonPropertyName("pollIntervalSeconds")]
    public int PollIntervalSeconds { get; set; } = 10;

    /// <summary>
    /// On-screen width, in pixels, of a grid tile. Its width:height ratio also sets the
    /// shape of the close-up each tile shows around the latest change. The close-up itself
    /// is rendered at twice this size, so the tile stays crisp and the full-screen view has
    /// plenty of detail to scale from.
    /// </summary>
    [JsonPropertyName("tileDisplayWidth")]
    public int TileDisplayWidth { get; set; } = 520;

    /// <summary>On-screen height, in pixels, of a grid tile. See <see cref="TileDisplayWidth"/>.</summary>
    [JsonPropertyName("tileDisplayHeight")]
    public int TileDisplayHeight { get; set; } = 390;

    /// <summary>
    /// How much of the page's width a tile's close-up covers, as a fraction (0.2 - 1.0).
    /// Smaller zooms in tighter - bigger text, less surrounding context. At the default 0.5
    /// ordinary typing is comfortably readable at tile size. The window grows automatically
    /// if the change itself is larger than this, such as a pasted screenshot.
    /// </summary>
    [JsonPropertyName("focusWindowFraction")]
    public double FocusWindowFraction { get; set; } = 0.5;

    /// <summary>
    /// How long, in seconds, a student's initials chip stays highlighted after a change is
    /// detected: it blinks red for most of that time (up to 3 seconds), then fades back to
    /// normal over the remainder. Short on purpose - a long flash makes everyone look
    /// active at once. At least 2.
    /// </summary>
    [JsonPropertyName("flashSeconds")]
    public int FlashSeconds { get; set; } = 5;

    /// <summary>Soft cap on students shown at once. The brief specifies 30.</summary>
    [JsonPropertyName("maxStudents")]
    public int MaxStudents { get; set; } = 30;

    /// <summary>
    /// Row label used when the picked notebook has no recognisable Class Notebook student
    /// structure (nothing to exclude), so the groups are just "section groups" rather than
    /// "students".
    /// </summary>
    [JsonPropertyName("genericGroupLabel")]
    public string GenericGroupLabel { get; set; } = "Section Group";

    /// <summary>The file the running config was loaded from, for display in diagnostics.</summary>
    [JsonIgnore]
    public string SourcePath { get; set; } = string.Empty;

    /// <summary>
    /// Loads <c>WorksheetWatcher.config.json</c> from the executable directory. A missing
    /// or malformed file is not fatal: sensible defaults are returned and the caller can
    /// surface a warning.
    /// </summary>
    /// <param name="warning">Populated with a human-readable problem, or null on success.</param>
    public static AppConfig Load(out string? warning)
    {
        warning = null;
        var path = Path.Combine(AppContext.BaseDirectory, "WorksheetWatcher.config.json");

        if (!File.Exists(path))
        {
            warning = $"Config file not found at {path}. Using built-in defaults.";
            return new AppConfig { SourcePath = path };
        }

        try
        {
            var json = File.ReadAllText(path);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                ReadCommentHandling = JsonCommentHandling.Skip,
                AllowTrailingCommas = true
            };
            var cfg = JsonSerializer.Deserialize<AppConfig>(json, options) ?? new AppConfig();
            cfg.SourcePath = path;
            if (cfg.MaxStudents < 1) cfg.MaxStudents = 30;
            if (cfg.TileDisplayWidth < 80) cfg.TileDisplayWidth = 520;
            if (cfg.TileDisplayHeight < 60) cfg.TileDisplayHeight = 390;
            cfg.FocusWindowFraction = Math.Clamp(cfg.FocusWindowFraction, 0.2, 1.0);
            if (cfg.FlashSeconds < 2) cfg.FlashSeconds = 5;
            return cfg;
        }
        catch (Exception ex)
        {
            warning = $"Config file at {path} could not be read ({ex.Message}). Using built-in defaults.";
            return new AppConfig { SourcePath = path };
        }
    }
}
