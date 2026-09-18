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
    /// <c>lastModifiedTime</c> before deciding whether to re-render it. Floored at 10
    /// seconds at run time regardless of what is configured, to protect OneNote's COM API
    /// from a misconfigured value.
    /// </summary>
    [JsonPropertyName("pollIntervalSeconds")]
    public int PollIntervalSeconds { get; set; } = 25;

    /// <summary>Base width, in pixels, a rendered page is rasterised to for a thumbnail tile.</summary>
    [JsonPropertyName("thumbnailWidth")]
    public int ThumbnailWidth { get; set; } = 320;

    /// <summary>Base height, in pixels, a rendered page is rasterised to for a thumbnail tile.</summary>
    [JsonPropertyName("thumbnailHeight")]
    public int ThumbnailHeight { get; set; } = 240;

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
            if (cfg.ThumbnailWidth < 40) cfg.ThumbnailWidth = 320;
            if (cfg.ThumbnailHeight < 30) cfg.ThumbnailHeight = 240;
            return cfg;
        }
        catch (Exception ex)
        {
            warning = $"Config file at {path} could not be read ({ex.Message}). Using built-in defaults.";
            return new AppConfig { SourcePath = path };
        }
    }
}
