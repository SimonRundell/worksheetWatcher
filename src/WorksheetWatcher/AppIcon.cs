using System.Drawing;

namespace WorksheetWatcher;

/// <summary>
/// Provides the application icon to every window. The icon is embedded as
/// <c>WorksheetWatcher.appicon.ico</c> and is also set as the executable icon via the
/// <c>ApplicationIcon</c> project property.
/// </summary>
internal static class AppIcon
{
    private static Icon? _cached;

    /// <summary>The shared application icon. Falls back to the system default if missing.</summary>
    public static Icon Value => _cached ??= Load();

    private static Icon Load()
    {
        try
        {
            var asm = typeof(AppIcon).Assembly;
            using var stream = asm.GetManifestResourceStream("WorksheetWatcher.appicon.ico");
            if (stream is not null)
                return new Icon(stream);
        }
        catch
        {
            // fall through to the system default
        }

        return SystemIcons.Application;
    }
}
