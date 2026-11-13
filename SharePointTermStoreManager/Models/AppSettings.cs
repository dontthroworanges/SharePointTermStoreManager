namespace SharePointTermStoreManager.Models;

/// <summary>Everything persisted to settings.json.</summary>
public sealed class AppSettings
{
    public List<ConnectionProfile> Profiles { get; set; } = new();
    public string? CurrentProfileId { get; set; }

    /// <summary>App (client) ID pre-filled in new profiles. Empty uses the built-in SharePoint Online client.</summary>
    public string DefaultClientId { get; set; } = "";

    /// <summary>Seconds before a success message closes itself; 0 keeps it until closed. Errors and warnings always stay.</summary>
    public int NotificationTimeoutSeconds { get; set; } = 5;

    /// <summary>"Default" (follow Windows), "Light" or "Dark".</summary>
    public string Theme { get; set; } = "Default";

    /// <summary>"#RRGGBB" from the Windows colours palette, or empty to use the Windows accent colour.</summary>
    public string AccentColor { get; set; } = "";

    /// <summary>Colourful icons in the tree and lists instead of outline glyphs.</summary>
    public bool UseColorfulIcons { get; set; }

    /// <summary>
    /// How many children the tree shows when an item is expanded, with "Show more" for the next page; 0 shows them
    /// all. Keeps large term sets quick to open (search finds items without scrolling to them).
    /// </summary>
    public int TreePageSize { get; set; } = 50;

    /// <summary>
    /// Whether the details pane's command bar has Delete. Off leaves deleting to the tree's right-click menu
    /// (and the Delete key), out of easy reach.
    /// </summary>
    public bool ShowDeleteInDetailsPane { get; set; } = true;

    /// <summary>Width of the term tree pane in pixels, as last dragged; 0 for the default 2:3 split with the editor.</summary>
    public double TreePaneWidth { get; set; }

    /// <summary>The Export page's last format (an <see cref="ExportFormat"/> name).</summary>
    public string ExportFormat { get; set; } = nameof(Models.ExportFormat.Excel);

    /// <summary>Whether the Export page includes deprecated terms.</summary>
    public bool ExportIncludeDeprecated { get; set; } = true;

    /// <summary>
    /// At startup, check GitHub for a newer release and download it in the background; the user is asked before
    /// it's installed.
    /// </summary>
    public bool CheckForUpdatesAtStartup { get; set; } = true;
}
