using System.ComponentModel;
using System.Reflection;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using SharePointTermStoreManager.Helpers;
using SharePointTermStoreManager.Models;
using SharePointTermStoreManager.Services;
using Windows.System;

namespace SharePointTermStoreManager.Views;

/// <summary>Settings: connections (profiles and sign-in), options and about. Changes are saved straight away.</summary>
public sealed partial class SettingsView : UserControl
{
    /// <summary>Width of the centred column of cards (as Windows 11 Settings), plus its padding.</summary>
    private const double MaxContentWidth = 1000 + 72;

    private bool _loading = true;
    private ConnectionProfile? _watchedProfile;

    public SettingsView()
    {
        InitializeComponent();

        LoadSettings();
        VersionCard.Header = $"SharePoint Term Store Manager {Assembly.GetExecutingAssembly().GetName().Version?.ToString(3)}";
        DataFolderCard.Description = SettingsStore.DataFolder;
        LogFolderCard.Description =
            $"A record of every change the app makes, with a copy of each item's details before it's merged or deleted (not including child terms). One JSON Lines file per month in {OperationLog.Folder}.";

        CategoryNav.SelectedItem = CategoryNav.MenuItems[0];
        App.Settings.CurrentProfileChanged += (_, _) => UpdateProfile();
        UpdateProfile();

        UpdateService.StateChanged += (_, _) => DispatcherQueue.TryEnqueue(UpdateUpdateCard);
        UpdateUpdateCard();

        StartMenuCard.Description =
            $"Add a shortcut to this app to your Start menu (for your account only, no admin rights needed). It follows the app when it updates itself. Runs {UpdateService.ExecutablePath}";
    }

    /// <summary>Shows the saved settings in the controls, without saving them again (also after a reset).</summary>
    public void LoadSettings()
    {
        _loading = true;
        var data = App.Settings.Data;
        TimeoutBox.Value = data.NotificationTimeoutSeconds;
        TreePageSizeBox.Value = data.TreePageSize;
        ShowDeleteSwitch.IsOn = data.ShowDeleteInDetailsPane;
        PickerScopeBox.SelectedIndex = data.PickerSearchScope == PickerScopes.WholeStore ? 1 : 0;
        ThemeBox.SelectedIndex = data.Theme switch { "Light" => 1, "Dark" => 2, _ => 0 };
        LoadAccentColor();
        IconStyleBox.SelectedIndex = data.UseColorfulIcons ? 1 : 0;
        DefaultClientIdBox.Text = data.DefaultClientId;
        AutoUpdateSwitch.IsOn = data.CheckForUpdatesAtStartup;
        StartMenuSwitch.IsOn = StartMenuShortcut.Exists;
        _loading = false;
    }

    /// <summary>Opens a category: "Connections", "Options" or "About".</summary>
    public void ShowCategory(string tag)
    {
        var item = CategoryNav.MenuItems.OfType<NavigationViewItem>().FirstOrDefault(i => (string)i.Tag == tag);
        if (item != null) CategoryNav.SelectedItem = item;
    }

    private void CategoryNav_SelectionChanged(NavigationView sender, NavigationViewSelectionChangedEventArgs args)
    {
        if (args.SelectedItem is not NavigationViewItem { Tag: string tag } item) return;
        CategoryTitle.Text = item.Content as string;
        ConnectionsPanel.Visibility = Vis(tag == "Connections");
        OptionsPanel.Visibility = Vis(tag == "Options");
        AppearancePanel.Visibility = Vis(tag == "Appearance");
        AboutPanel.Visibility = Vis(tag == "About");
    }

    private static Visibility Vis(bool visible) => visible ? Visibility.Visible : Visibility.Collapsed;

    /// <summary>
    /// Like Windows 11 Settings: the column of cards is capped at <see cref="MaxContentWidth"/> and centred when
    /// there's more room. Set from the viewport's width so nothing depends on the panel being re-measured.
    /// </summary>
    private void ContentScroller_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        double width = e.NewSize.Width;
        ContentPanel.Width = Math.Min(width, MaxContentWidth);
        ContentPanel.Margin = new Thickness(Math.Max(0, Math.Floor((width - MaxContentWidth) / 2)), 0, 0, 0);
    }

    // ---- Connections ---------------------------------------------------------------------------------------

    private async void EditProfile_Click(object sender, RoutedEventArgs e)
    {
        if (App.Settings.CurrentProfile is { } profile) await App.MainWindow.EditProfileAsync(profile);
    }

    private async void AddProfile_Click(object sender, RoutedEventArgs e) => await App.MainWindow.EditProfileAsync(null);

    private async void SignOut_Click(object sender, RoutedEventArgs e) => await App.MainWindow.SignOutAsync();

    private void DefaultClientIdBox_LostFocus(object sender, RoutedEventArgs e)
    {
        var clientId = DefaultClientIdBox.Text.Trim();
        if (clientId == App.Settings.Data.DefaultClientId) return;
        App.Settings.Data.DefaultClientId = clientId;
        App.Settings.Save();
    }

    private void UpdateProfile()
    {
        var profile = App.Settings.CurrentProfile;
        WatchProfile(profile);

        ProfilePicture.DisplayName = profile?.EffectiveName ?? "";
        ProfileNameText.Text = profile?.EffectiveName ?? "No profile";
        ProfileAccountText.Text = profile?.AccountDisplay ?? "Add a SharePoint site to get started";

        CurrentProfileCard.Header = profile?.EffectiveName ?? "No profile yet";
        CurrentProfileCard.Description = profile == null
            ? "Add a SharePoint site to open its term store."
            : $"{profile.SiteUrl}\n{profile.AccountDisplay}";
        EditProfileButton.IsEnabled = profile != null;
        SignOutButton.Visibility = Vis(profile?.IsSignedIn == true);
    }

    /// <summary>Follows the current profile so the header updates the moment sign-in state changes.</summary>
    private void WatchProfile(ConnectionProfile? profile)
    {
        if (_watchedProfile == profile) return;
        if (_watchedProfile != null) _watchedProfile.PropertyChanged -= WatchedProfile_PropertyChanged;
        _watchedProfile = profile;
        if (profile != null) profile.PropertyChanged += WatchedProfile_PropertyChanged;
    }

    private void WatchedProfile_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (DispatcherQueue.HasThreadAccess) UpdateProfile();
        else DispatcherQueue.TryEnqueue(UpdateProfile);
    }

    // ---- Options -------------------------------------------------------------------------------------------

    private void TimeoutBox_ValueChanged(NumberBox sender, NumberBoxValueChangedEventArgs args)
    {
        if (_loading) return;
        var seconds = double.IsNaN(sender.Value) ? 0 : sender.Value;
        App.Settings.Data.NotificationTimeoutSeconds = (int)Math.Clamp(Math.Round(seconds), 0, 120);
        App.Settings.Save();
    }

    private void TreePageSizeBox_ValueChanged(NumberBox sender, NumberBoxValueChangedEventArgs args)
    {
        if (_loading) return;
        if (double.IsNaN(sender.Value))
        {
            // Cleared: keep the current value rather than reading it as 0 (show everything).
            sender.Value = App.Settings.Data.TreePageSize;
            return;
        }
        App.Settings.Data.TreePageSize = (int)Math.Clamp(Math.Round(sender.Value), 0, 5000);
        App.Settings.Save();
    }

    private void ShowDeleteSwitch_Toggled(object sender, RoutedEventArgs e)
    {
        if (_loading) return;
        App.Settings.Data.ShowDeleteInDetailsPane = ShowDeleteSwitch.IsOn;
        App.Settings.Save();
        App.MainWindow.RefreshEditorCommands();
    }

    private void PickerScopeBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading || PickerScopeBox.SelectedItem is not ComboBoxItem { Tag: string scope }) return;
        App.Settings.Data.PickerSearchScope = scope;
        App.Settings.Save();
    }

    private async void OpenLogFolder_Click(object sender, RoutedEventArgs e)
    {
        Directory.CreateDirectory(OperationLog.Folder);
        await Launcher.LaunchFolderPathAsync(OperationLog.Folder);
    }

    // ---- Appearance ----------------------------------------------------------------------------------------

    private void ThemeBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading || ThemeBox.SelectedItem is not ComboBoxItem { Tag: string theme }) return;
        App.Settings.Data.Theme = theme;
        App.MainWindow.ApplyTheme(theme);
        App.Settings.Save();
    }

    private void LoadAccentColor()
    {
        var saved = App.Settings.Data.AccentColor;
        bool custom = AccentColors.Parse(saved) != null;
        AccentModeBox.SelectedIndex = custom ? 1 : 0;
        AccentPalettePanel.Visibility = Vis(custom);
        AccentGrid.SelectedItem = AccentColors.Palette.FirstOrDefault(s => string.Equals(s.Hex, saved, StringComparison.OrdinalIgnoreCase));
    }

    private void AccentModeBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading) return;
        bool custom = AccentModeBox.SelectedIndex == 1;
        AccentPalettePanel.Visibility = Vis(custom);
        if (!custom)
        {
            SetAccentColor("");
            return;
        }
        // Start from the Windows accent colour when it's one of the swatches, as Windows Settings does.
        var system = new Windows.UI.ViewManagement.UISettings().GetColorValue(Windows.UI.ViewManagement.UIColorType.Accent);
        var swatch = AccentGrid.SelectedItem as AccentSwatch
            ?? AccentColors.Palette.FirstOrDefault(s => s.Color == system)
            ?? AccentColors.Palette.First(s => s.Hex == "#0078D4");
        AccentGrid.SelectedItem = swatch;
        SetAccentColor(swatch.Hex);
    }

    private void AccentGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading || AccentGrid.SelectedItem is not AccentSwatch swatch) return;
        SetAccentColor(swatch.Hex);
    }

    private static void SetAccentColor(string hex)
    {
        if (string.Equals(App.Settings.Data.AccentColor, hex, StringComparison.OrdinalIgnoreCase)) return;
        App.Settings.Data.AccentColor = hex;
        App.MainWindow.ApplyAccentColor(hex);
        App.Settings.Save();
    }

    private void IconStyleBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loading) return;
        bool colorful = IconStyleBox.SelectedIndex == 1;
        App.Settings.Data.UseColorfulIcons = colorful;
        DisplayOptions.Current.UseColorfulIcons = colorful;
        App.Settings.Save();
    }

    // ---- About ---------------------------------------------------------------------------------------------

    private async void OpenDataFolder_Click(object sender, RoutedEventArgs e)
    {
        Directory.CreateDirectory(SettingsStore.DataFolder);
        await Launcher.LaunchFolderPathAsync(SettingsStore.DataFolder);
    }

    /// <summary>Shows the update check's progress and result, whether it started here or at startup.</summary>
    private void UpdateUpdateCard()
    {
        string current = UpdateService.CurrentVersion.ToString(3);
        string? available = UpdateService.AvailableVersion?.ToString(3);
        var state = UpdateService.State;
        bool working = state is UpdateState.Checking or UpdateState.Downloading;

        UpdateCard.Description = state switch
        {
            UpdateState.Checking => "Checking GitHub for a newer version…",
            UpdateState.Downloading => $"Downloading version {available}… {UpdateService.DownloadPercent}%",
            UpdateState.Ready => $"Version {available} is downloaded and ready to install. You have {current}.",
            UpdateState.UpToDate => $"You're up to date (version {current}).",
            UpdateState.Failed => UpdateService.LastError ?? "Couldn't check for updates.",
            _ => $"You have version {current}. Newer releases are downloaded from GitHub and replace this app's .exe.",
        };
        UpdateButton.Content = state == UpdateState.Ready ? "Install and restart" : "Check for updates";
        UpdateButton.IsEnabled = !working;
        UpdateProgressRing.IsActive = working;
        UpdateProgressRing.Visibility = Vis(working);
    }

    private async void UpdateButton_Click(object sender, RoutedEventArgs e)
    {
        var update = UpdateService.ReadyUpdate;
        if (UpdateService.State != UpdateState.Ready || update == null)
        {
            try
            {
                update = await UpdateService.CheckAndDownloadAsync();
            }
            catch (UpdateException)
            {
                return; // shown on the card
            }
        }
        if (update != null) await App.MainWindow.OfferUpdateAsync(update);
    }

    private void StartMenuSwitch_Toggled(object sender, RoutedEventArgs e)
    {
        if (_loading) return;
        try
        {
            if (StartMenuSwitch.IsOn) StartMenuShortcut.Create();
            else StartMenuShortcut.Remove();
        }
        catch (Exception ex)
        {
            App.MainWindow.ShowStatus($"Couldn't {(StartMenuSwitch.IsOn ? "add" : "remove")} the Start menu shortcut: {ex.Message}", InfoBarSeverity.Error);
            _loading = true;
            StartMenuSwitch.IsOn = StartMenuShortcut.Exists;
            _loading = false;
        }
    }

    private void AutoUpdateSwitch_Toggled(object sender, RoutedEventArgs e)
    {
        if (_loading) return;
        App.Settings.Data.CheckForUpdatesAtStartup = AutoUpdateSwitch.IsOn;
        App.Settings.Save();
    }

    /// <summary>
    /// Reset app to defaults, as in SharePoint Search Explorer. The operations log is kept unless asked: it's the
    /// record of changes made to the term store, with copies of merged and deleted items.
    /// </summary>
    private async void ResetApp_Click(object sender, RoutedEventArgs e)
    {
        var deleteLog = new CheckBox { Content = "Also delete the operations log" };
        var confirm = new ContentDialog
        {
            XamlRoot = XamlRoot,
            RequestedTheme = App.MainWindow.DialogTheme,
            Title = "Reset app to defaults?",
            Content = new StackPanel
            {
                Spacing = 12,
                Children =
                {
                    new TextBlock
                    {
                        Text = "All profiles will be removed and signed out, the sign-in window's saved session will be cleared and every setting will return to its default. Nothing in SharePoint is changed. This can't be undone.",
                        TextWrapping = TextWrapping.Wrap,
                    },
                    deleteLog,
                },
            },
            PrimaryButtonText = "Reset",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Close,
        };
        if (await confirm.ShowAsync() != ContentDialogResult.Primary) return;
        await App.MainWindow.ResetToDefaultsAsync(deleteLog.IsChecked == true);
    }
}
