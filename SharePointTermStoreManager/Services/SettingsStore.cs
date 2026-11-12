using System.Diagnostics;
using System.Text.Json;
using SharePointTermStoreManager.Models;

namespace SharePointTermStoreManager.Services;

/// <summary>Loads and saves <see cref="AppSettings"/> as JSON under %LOCALAPPDATA%\SharePointTermStoreManager.</summary>
public sealed class SettingsStore
{
    /// <summary>Settings, the sign-in window's browser profile and the encrypted token cache live here.</summary>
    public static string DataFolder { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SharePointTermStoreManager");

    private static string SettingsPath => Path.Combine(DataFolder, "settings.json");

    private static readonly JsonSerializerOptions s_json = new() { WriteIndented = true };

    private SettingsStore(AppSettings data) => Data = data;

    public AppSettings Data { get; }

    public event EventHandler? CurrentProfileChanged;

    public ConnectionProfile? CurrentProfile =>
        Data.Profiles.FirstOrDefault(p => p.Id == Data.CurrentProfileId) ?? Data.Profiles.FirstOrDefault();

    public static SettingsStore Load()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var data = JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(SettingsPath), s_json);
                if (data != null) return new SettingsStore(data);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Could not read settings, starting fresh: {ex.Message}");
        }
        return new SettingsStore(new AppSettings());
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(DataFolder);
            var temp = SettingsPath + ".tmp";
            File.WriteAllText(temp, JsonSerializer.Serialize(Data, s_json));
            File.Move(temp, SettingsPath, overwrite: true);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Could not save settings: {ex.Message}");
        }
    }

    public void SetCurrentProfile(ConnectionProfile profile)
    {
        Data.CurrentProfileId = profile.Id;
        Save();
        CurrentProfileChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Adds a new profile or saves changes to an existing one.</summary>
    public void SaveProfile(ConnectionProfile profile)
    {
        if (!Data.Profiles.Contains(profile)) Data.Profiles.Add(profile);
        Save();
        if (profile == CurrentProfile) CurrentProfileChanged?.Invoke(this, EventArgs.Empty);
    }

    public void RemoveProfile(ConnectionProfile profile)
    {
        bool wasCurrent = profile == CurrentProfile;
        Data.Profiles.Remove(profile);
        if (wasCurrent) Data.CurrentProfileId = Data.Profiles.FirstOrDefault()?.Id;
        Save();
        if (wasCurrent) CurrentProfileChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Restores every setting to its default and removes all profiles, as in SharePoint Search Explorer. Values
    /// are copied into the existing <see cref="Data"/> instance because the app holds on to it.
    /// </summary>
    public void ResetToDefaults()
    {
        var defaults = new AppSettings();
        foreach (var property in typeof(AppSettings).GetProperties().Where(p => p.CanRead && p.CanWrite))
            property.SetValue(Data, property.GetValue(defaults));
        Save();
        CurrentProfileChanged?.Invoke(this, EventArgs.Empty);
    }
}
