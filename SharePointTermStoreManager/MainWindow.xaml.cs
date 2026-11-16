using System.Collections.ObjectModel;
using Microsoft.SharePoint.Client;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Automation.Peers;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using SharePointTermStoreManager.Helpers;
using SharePointTermStoreManager.Models;
using SharePointTermStoreManager.Services;
using SharePointTermStoreManager.Views;
using Windows.System;

namespace SharePointTermStoreManager;

public sealed partial class MainWindow : Window
{
    /// <summary>Term groups at the top of the tree.</summary>
    public ObservableCollection<TermNode> Roots { get; } = new();

    private CancellationTokenSource? _loadCts;
    private CancellationTokenSource? _detailsCts;

    /// <summary>True once a term store is shown.</summary>
    private bool _storeLoaded;

    /// <summary>True while a change is being sent to SharePoint (or a load the change depends on).</summary>
    private bool _busy;

    /// <summary>The node shown in the editor.</summary>
    private TermNode? _selected;

    private bool _closeConfirmed;

    /// <summary>Closes success messages after the time set in Settings.</summary>
    private readonly Microsoft.UI.Dispatching.DispatcherQueueTimer _statusTimer;

    private CancellationTokenSource? _searchCts;

    public MainWindow()
    {
        InitializeComponent();

        ExtendsContentIntoTitleBar = true;
        SetTitleBar(AppTitleBar);
        AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico"));
        AppWindow.Resize(new Windows.Graphics.SizeInt32(1280, 820));
        if (AppWindow.Presenter is OverlappedPresenter presenter)
        {
            presenter.PreferredMinimumWidth = 820;
            presenter.PreferredMinimumHeight = 520;
        }
        AppWindow.Closing += AppWindow_Closing;
        RootGrid.SizeChanged += (_, e) =>
            // TitleBar sizes its content to fit, so give the search bar an explicit share of the window width
            // (after the icon and app name, the caption buttons and drag space).
            SearchBar.Width = Math.Clamp(e.NewSize.Width - 520, 260, 820);
        RootGrid.Loaded += RootGrid_Loaded;

        ApplyTheme(App.Settings.Data.Theme);
        // While following Windows, pick up an accent colour changed in Windows Settings.
        AccentColors.SystemAccentChanged += (_, _) => DispatcherQueue.TryEnqueue(() =>
        {
            if (AccentColors.UsesSystemAccent) ApplyAccentColor("");
        });

        _statusTimer = DispatcherQueue.CreateTimer();
        _statusTimer.IsRepeating = false;
        _statusTimer.Tick += (_, _) => StatusBar.IsOpen = false;

        Editor.SaveRequested += async (_, _) => await SaveAsync();
        Editor.SetPlaceholderCommands(PlaceholderCommands);
        // Clicking a parent in the editor's breadcrumb selects it, as clicking it in the tree would.
        Editor.PathItemInvoked += async (_, node) => await SelectAsync(node);
        App.Settings.CurrentProfileChanged += (_, _) => UpdateProfileDisplay();

        NavView.SelectedItem = TermStoreNavItem;
        UpdateProfileDisplay();
        UpdateCommands();
        HookTreeSplitter();
        HookSearchPassthrough();
    }

    // ---- Title bar: the search box isn't part of the window's drag area ---------------------------------------

    /// <summary>The passthrough rectangles last added for the search box and button (physical pixels).</summary>
    private Windows.Graphics.RectInt32[] _searchPassthrough = [];

    private bool _passthroughPending;

    /// <summary>
    /// The search box sits in the title bar, where a press starts a window drag unless the spot is marked as
    /// passthrough. TitleBar marks its content when the title bar resizes, but the search bar's width is set after
    /// that (and follows the window width), so the marked area can miss parts of the box. After each layout pass,
    /// make sure the search box and button are marked, once the TitleBar has made its own update.
    /// </summary>
    private void HookSearchPassthrough()
    {
        SearchBar.LayoutUpdated += (_, _) =>
        {
            if (_passthroughPending) return;
            _passthroughPending = true;
            DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
            {
                _passthroughPending = false;
                UpdateSearchPassthrough();
            });
        };
    }

    private void UpdateSearchPassthrough()
    {
        if (SearchBar.XamlRoot is not { } root || SearchBar.ActualWidth <= 0) return;
        var wanted = new[] { PhysicalBounds(GlobalSearchBox, root), PhysicalBounds(GlobalSearchButton, root) };

        var source = Microsoft.UI.Input.InputNonClientPointerSource.GetForWindowId(AppWindow.Id);
        var current = source.GetRegionRects(Microsoft.UI.Input.NonClientRegionKind.Passthrough) ?? [];
        if (wanted.All(current.Contains)) return;

        // Keep the TitleBar's own rectangles; replace the ones added here last time (the box may have moved).
        var rects = current.Where(r => !_searchPassthrough.Contains(r)).Concat(wanted).Distinct().ToArray();
        source.SetRegionRects(Microsoft.UI.Input.NonClientRegionKind.Passthrough, rects);
        _searchPassthrough = wanted;
    }

    /// <summary>An element's bounds in the window's client area, in physical pixels.</summary>
    private static Windows.Graphics.RectInt32 PhysicalBounds(FrameworkElement element, XamlRoot root)
    {
        double scale = root.RasterizationScale;
        var b = element.TransformToVisual(null)
            .TransformBounds(new Windows.Foundation.Rect(0, 0, element.ActualWidth, element.ActualHeight));
        int x = (int)Math.Floor(b.X * scale), y = (int)Math.Floor(b.Y * scale);
        return new Windows.Graphics.RectInt32(x, y,
            (int)Math.Ceiling(b.Right * scale) - x, (int)Math.Ceiling(b.Bottom * scale) - y);
    }

    // ---- Tree / editor splitter ----------------------------------------------------------------------------

    /// <summary>The splitter's column (the gap between the cards).</summary>
    private const double SplitterWidth = 12;

    private double _treeWidthAtDragStart;

    private void HookTreeSplitter()
    {
        TreeSplitter.DragStarted += (_, _) => _treeWidthAtDragStart = TreeColumn.ActualWidth;
        TreeSplitter.DragDelta += (_, delta) => ApplyTreeWidth(_treeWidthAtDragStart + delta);
        TreeSplitter.Nudged += (_, delta) => ApplyTreeWidth(TreeColumn.ActualWidth + delta);
        TreeSplitter.DragCompleted += (_, _) =>
        {
            App.Settings.Data.TreePaneWidth = Math.Round(TreeColumn.ActualWidth);
            App.Settings.Save();
        };
        TreeSplitter.ResetRequested += (_, _) =>
        {
            ApplyDefaultSplit();
            App.Settings.Data.TreePaneWidth = 0;
            App.Settings.Save();
        };
        // Keep the remembered width, but never squeeze the editor below its minimum when the window is narrower.
        PanesGrid.SizeChanged += (_, _) =>
        {
            if (_panesSplit && App.Settings.Data.TreePaneWidth > 0) ApplyTreeWidth(App.Settings.Data.TreePaneWidth);
        };
    }

    /// <summary>The editor's minimum width while it's shown.</summary>
    private const double EditorMinWidth = 360;

    /// <summary>True while the tree and editor are side by side (a term store is loaded).</summary>
    private bool _panesSplit;

    /// <summary>
    /// Until a term store is loaded the editor has nothing to show, so the tree card (with its connect prompt)
    /// takes the whole width and the editor, splitter and the tree's header bar are hidden. Once loaded, the
    /// panes go back to the remembered split.
    /// </summary>
    private void UpdatePaneLayout(bool split)
    {
        if (split == _panesSplit) return;
        _panesSplit = split;

        var visibility = split ? Visibility.Visible : Visibility.Collapsed;
        EditorCard.Visibility = TreeSplitter.Visibility = TreeHeader.Visibility = visibility;
        SplitterColumn.Width = new GridLength(split ? SplitterWidth : 0);
        EditorColumn.MinWidth = split ? EditorMinWidth : 0;

        if (!split)
        {
            TreeColumn.Width = new GridLength(1, GridUnitType.Star);
            EditorColumn.Width = new GridLength(0);
        }
        else if (App.Settings.Data.TreePaneWidth > 0)
        {
            ApplyTreeWidth(App.Settings.Data.TreePaneWidth);
        }
        else
        {
            ApplyDefaultSplit();
        }
    }

    /// <summary>The default 2:3 split between the tree and the editor.</summary>
    private void ApplyDefaultSplit()
    {
        TreeColumn.Width = new GridLength(2, GridUnitType.Star);
        EditorColumn.Width = new GridLength(3, GridUnitType.Star);
    }

    /// <summary>
    /// When the tree pane is narrow, its command bar shows icons only (the tooltips still name them), so the
    /// profile name beside it keeps its room. AppBarButton's style fixes its width at 68px (room for a label
    /// underneath), which only label-on-right mode overrides, so icon-only buttons get a square 40px width
    /// (the bar's compact height) instead of keeping the wide spacing.
    /// </summary>
    private void TreeCard_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        bool iconsOnly = e.NewSize.Width < 440;
        TreeCommands.DefaultLabelPosition = iconsOnly
            ? CommandBarDefaultLabelPosition.Collapsed
            : CommandBarDefaultLabelPosition.Right;
        foreach (var command in TreeCommands.PrimaryCommands)
            if (command is AppBarButton button) button.Width = iconsOnly ? 40 : double.NaN;
    }

    /// <summary>Sets the tree pane to <paramref name="width"/> pixels, within both panes' minimum widths.</summary>
    private void ApplyTreeWidth(double width)
    {
        double available = PanesGrid.ActualWidth;
        if (available <= 0) return;
        double max = Math.Max(TreeColumn.MinWidth, available - SplitterWidth - EditorColumn.MinWidth);
        TreeColumn.Width = new GridLength(Math.Clamp(width, TreeColumn.MinWidth, max));
        EditorColumn.Width = new GridLength(1, GridUnitType.Star);
    }

    private ConnectionProfile? CurrentProfile => App.Settings.CurrentProfile;

    private async void RootGrid_Loaded(object sender, RoutedEventArgs e)
    {
        if (App.Settings.Data.CheckForUpdatesAtStartup) _ = CheckForUpdatesAtStartupAsync();
        // First run: ask for a site to connect to.
        if (CurrentProfile == null) await EditProfileAsync(null);
    }

    // ---- Updates -------------------------------------------------------------------------------------------

    /// <summary>
    /// Checks GitHub in the background; a newer version is downloaded, then offered. Quiet when up to date or
    /// offline: a failed check shows in Settings › About, not as an error over the term store.
    /// </summary>
    private async Task CheckForUpdatesAtStartupAsync()
    {
        UpdateInfo? update;
        try
        {
            update = await UpdateService.CheckAndDownloadAsync();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Update check failed: {ex.Message}");
            return;
        }
        if (update != null) await OfferUpdateAsync(update, fromStartup: true);
    }

    /// <summary>
    /// Asks whether to install a downloaded update now. Installing replaces the .exe the app runs from and
    /// restarts it; unsaved changes in the details pane are offered for saving first.
    /// </summary>
    public async Task OfferUpdateAsync(UpdateInfo update, bool fromStartup = false)
    {
        string version = update.Version.ToString(3);
        var content = new StackPanel { Spacing = 12 };
        content.Children.Add(new TextBlock
        {
            Text = $"Version {version} has been downloaded (you have {UpdateService.CurrentVersion.ToString(3)}). " +
                   "The app will close, replace its .exe with the new version and start again.",
            TextWrapping = TextWrapping.Wrap,
        });
        if (!string.IsNullOrWhiteSpace(update.Notes))
        {
            content.Children.Add(new TextBlock { Text = "What's new", Style = (Style)Application.Current.Resources["BodyStrongTextBlockStyle"] });
            content.Children.Add(new ScrollViewer
            {
                MaxHeight = 220,
                Content = new TextBlock { Text = update.Notes.Trim(), TextWrapping = TextWrapping.Wrap, IsTextSelectionEnabled = true },
            });
        }
        content.Children.Add(new HyperlinkButton
        {
            Content = "View the release on GitHub",
            NavigateUri = new Uri(update.ReleaseUrl),
            Padding = new Thickness(0),
        });

        var dialog = new ContentDialog
        {
            XamlRoot = Content.XamlRoot,
            RequestedTheme = DialogTheme,
            Title = $"Update to version {version}?",
            Content = content,
            PrimaryButtonText = "Restart and update",
            CloseButtonText = "Later",
            DefaultButton = ContentDialogButton.Primary,
        };

        ContentDialogResult result;
        try
        {
            result = await dialog.ShowAsync();
        }
        catch (Exception)
        {
            // Another dialog is open (e.g. the first-run profile dialog): leave a note instead of interrupting it.
            if (fromStartup)
                ShowStatus($"Version {version} is ready to install. Go to Settings › About to update.", InfoBarSeverity.Informational);
            return;
        }
        if (result != ContentDialogResult.Primary) return;
        await InstallUpdateAsync(update);
    }

    private async Task InstallUpdateAsync(UpdateInfo update)
    {
        if (_busy)
        {
            ShowStatus("Wait for the current change to finish, then install the update.", InfoBarSeverity.Warning);
            return;
        }
        if (!await ConfirmLeaveAsync()) return;

        try
        {
            if (!UpdateService.Install(update))
            {
                ShowStatus("The update wasn't installed: Windows needs administrator permission to replace the app in its folder.", InfoBarSeverity.Warning);
                return;
            }
        }
        catch (Exception ex)
        {
            ShowError("Couldn't install the update", ex);
            return;
        }

        // The new version has started; close this one without asking again.
        _closeConfirmed = true;
        Application.Current.Exit();
    }

    // ---- Navigation ----------------------------------------------------------------------------------------

    private void NavView_ItemInvoked(NavigationView sender, NavigationViewItemInvokedEventArgs args)
    {
        switch ((args.InvokedItemContainer as NavigationViewItem)?.Tag as string)
        {
            case "termstore":
                ShowTermStore();
                break;
            case "log":
                ShowOperationsLog();
                break;
            case "export":
                ShowExport();
                break;
            case "profile":
                ShowProfileMenu();
                break;
            case "settings":
                ShowSettings();
                break;
        }
    }

    /// <summary>
    /// Shows the term store page. The log and Settings are separate views, so the tree and any edits stay as they
    /// were.
    /// </summary>
    public void ShowTermStore() => ShowPage(TermStoreView, TermStoreNavItem);

    /// <summary>Shows the operations log, reading the log files again if something was logged since.</summary>
    public void ShowOperationsLog()
    {
        ShowPage(LogView, LogNavItem);
        LogView.Activate();
    }

    /// <summary>Shows Settings, optionally at a category ("Connections", "Options" or "About").</summary>
    public void ShowSettings(string? category = null)
    {
        ShowPage(SettingsView, SettingsNavItem);
        if (category != null) SettingsView.ShowCategory(category);
    }

    /// <summary>Shows the Export page.</summary>
    public void ShowExport()
    {
        ShowPage(ExportView, ExportNavItem);
        ExportView.UpdateState(_storeLoaded, Roots);
    }

    private void ShowPage(UIElement page, NavigationViewItem item)
    {
        NavView.SelectedItem = item;
        foreach (var view in new UIElement[] { TermStoreView, LogView, ExportView, SettingsView })
            view.Visibility = view == page ? Visibility.Visible : Visibility.Collapsed;
    }

    // ---- Appearance ----------------------------------------------------------------------------------------

    /// <summary>
    /// The theme for dialogs and other windows: the theme chosen in Settings, or Default (the app's theme, which
    /// follows Windows) when none is chosen, matching the window. Worked out from the setting rather than
    /// RootGrid.ActualTheme, because <see cref="AccentColors.Refresh"/> briefly flips the window to the opposite
    /// theme to repaint the accent colour, and a dialog opened in that moment (e.g. right after Reset) would
    /// get the wrong one.
    /// </summary>
    public ElementTheme DialogTheme => App.Settings.Data.Theme switch
    {
        "Light" => ElementTheme.Light,
        "Dark" => ElementTheme.Dark,
        _ => ElementTheme.Default,
    };

    /// <summary>"Default" (follow Windows), "Light" or "Dark", for the window and its title bar buttons.</summary>
    public void ApplyTheme(string theme)
    {
        RootGrid.RequestedTheme = theme switch
        {
            "Light" => ElementTheme.Light,
            "Dark" => ElementTheme.Dark,
            _ => ElementTheme.Default,
        };
        AppWindow.TitleBar.PreferredTheme = theme switch
        {
            "Light" => TitleBarTheme.Light,
            "Dark" => TitleBarTheme.Dark,
            _ => TitleBarTheme.UseDefaultAppMode,
        };
    }

    /// <summary>Applies an accent colour ("#RRGGBB", or empty for the Windows accent colour).</summary>
    public void ApplyAccentColor(string hex)
    {
        AccentColors.Apply(hex);
        AccentColors.Refresh(RootGrid);
    }

    // ---- Profiles ------------------------------------------------------------------------------------------

    private ConnectionProfile? _watchedProfile;
    private bool _profileDialogOpen;

    private void UpdateProfileDisplay()
    {
        var profile = CurrentProfile;
        WatchProfile(profile);

        ProfileNavItem.Content = profile?.EffectiveName ?? "Add a profile";
        TreeProfileText.Text = EmptyStateProfileText.Text = profile?.EffectiveName ?? "";
        EmptyStateProfileText.Visibility = profile == null ? Visibility.Collapsed : Visibility.Visible;
        ToolTipService.SetToolTip(TreeProfileText, profile?.SiteUrl);
        ToolTipService.SetToolTip(EmptyStateProfileText, profile?.SiteUrl);
        ToolTipService.SetToolTip(ProfileNavItem, profile == null
            ? "Add a SharePoint site profile"
            : $"{profile.EffectiveName}\n{profile.SiteUrl}\n{profile.AccountDisplay}");
        Title = profile == null ? "SharePoint Term Store Manager" : $"SharePoint Term Store Manager – {profile.EffectiveName}";

        // Green check when signed in, yellow alert otherwise.
        bool signedIn = profile?.IsSignedIn == true;
        ApplyStatusBadge(ProfileBadge, signedIn);
        AutomationProperties.SetName(ProfileNavItem, signedIn ? "Profile, signed in" : "Profile, not signed in");
        UpdateCommands();
    }

    /// <summary>Follows the current profile so the badge updates the moment sign-in state changes.</summary>
    private void WatchProfile(ConnectionProfile? profile)
    {
        if (_watchedProfile == profile) return;
        if (_watchedProfile != null) _watchedProfile.PropertyChanged -= WatchedProfile_PropertyChanged;
        _watchedProfile = profile;
        if (profile != null) profile.PropertyChanged += WatchedProfile_PropertyChanged;
    }

    private void WatchedProfile_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName is not (nameof(ConnectionProfile.IsSignedIn) or nameof(ConnectionProfile.UserName) or nameof(ConnectionProfile.Name))) return;
        if (DispatcherQueue.HasThreadAccess) UpdateProfileDisplay();
        else DispatcherQueue.TryEnqueue(UpdateProfileDisplay);
    }

    /// <summary>
    /// Profile switcher: the current profile opens a submenu with its account actions; other profiles switch
    /// (and connect) on click. "Add profile" comes last.
    /// </summary>
    private void ShowProfileMenu()
    {
        var settings = App.Settings;
        var current = settings.CurrentProfile;
        var menu = new MenuFlyout();

        foreach (var profile in settings.Data.Profiles)
        {
            if (profile == current)
            {
                menu.Items.Add(CurrentProfileMenuItem(profile));
                continue;
            }
            var item = MenuItem(profile.EffectiveName, FluentIcons.City, async () => await SwitchProfileAsync(profile));
            ToolTipService.SetToolTip(item, $"Switch to this profile\n{profile.SiteUrl}");
            menu.Items.Add(item);
        }
        if (settings.Data.Profiles.Count > 0) menu.Items.Add(new MenuFlyoutSeparator());

        menu.Items.Add(MenuItem("Add profile…", FluentIcons.Add, async () => await EditProfileAsync(null)));

        menu.ShowAt(ProfileNavItem, new FlyoutShowOptions { Placement = FlyoutPlacementMode.RightEdgeAlignedBottom });
    }

    private MenuFlyoutSubItem CurrentProfileMenuItem(ConnectionProfile profile)
    {
        var sub = new MenuFlyoutSubItem { Text = profile.EffectiveName, Icon = FluentIcons.Create(FluentIcons.City) };
        ToolTipService.SetToolTip(sub, $"Current profile\n{profile.SiteUrl}");
        AutomationProperties.SetName(sub, $"{profile.EffectiveName}, current profile");
        AddCurrentProfileBadge(sub);

        sub.Items.Add(new MenuFlyoutItem
        {
            Text = profile.AccountDisplay,
            Icon = FluentIcons.Create(profile.IsSignedIn ? FluentIcons.CheckmarkCircle : FluentIcons.Warning),
            IsEnabled = false,
        });
        sub.Items.Add(_storeLoaded
            ? MenuItem("Refresh term store", FluentIcons.ArrowClockwise, async () => await RefreshAsync())
            : MenuItem("Connect", FluentIcons.PersonArrowRight, async () => await RefreshAsync()));
        sub.Items.Add(MenuItem("Sign in with a different account…", FluentIcons.PersonSwap, async () => await ChangeAccountAsync()));
        sub.Items.Add(MenuItem("Edit profile…", FluentIcons.Edit, async () => await EditProfileAsync(profile)));
        if (profile.IsSignedIn) sub.Items.Add(MenuItem("Sign out", FluentIcons.SignOut, async () => await SignOutAsync()));

        sub.Items.Add(new MenuFlyoutSeparator());
        sub.Items.Add(MenuItem("Remove profile", FluentIcons.Delete, async () => await RemoveProfileAsync(profile)));
        return sub;
    }

    private static MenuFlyoutItem MenuItem(string text, string glyph, Action onClick)
    {
        var item = new MenuFlyoutItem { Text = text, Icon = FluentIcons.Create(glyph) };
        item.Click += (_, _) => onClick();
        return item;
    }

    /// <summary>
    /// Puts the green check badge to the right of the current profile's name. MenuFlyoutSubItem only takes text,
    /// so the badge is added to its template, just before the submenu chevron.
    /// </summary>
    private static void AddCurrentProfileBadge(MenuFlyoutSubItem item)
    {
        item.Loaded += (_, _) =>
        {
            if (FindDescendant(item, "SubItemChevron") is not { Parent: Grid grid } chevron) return;
            if (grid.Children.OfType<InfoBadge>().Any()) return;

            var badge = CreateStatusBadge(signedIn: true);
            badge.Margin = new Thickness(12, 0, 0, 0);
            badge.HorizontalAlignment = HorizontalAlignment.Left;
            badge.VerticalAlignment = VerticalAlignment.Center;
            AutomationProperties.SetAccessibilityView(badge, AccessibilityView.Raw);
            Grid.SetColumn(badge, Grid.GetColumn(chevron));
            grid.Children.Add(badge);

            var m = chevron.Margin;
            chevron.Margin = new Thickness(Math.Max(m.Left, 12 + badge.Width + 8), m.Top, m.Right, m.Bottom);
        };
    }

    private static FrameworkElement? FindDescendant(DependencyObject root, string name)
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(root); i++)
        {
            var child = VisualTreeHelper.GetChild(root, i);
            if (child is FrameworkElement fe && fe.Name == name) return fe;
            if (FindDescendant(child, name) is { } found) return found;
        }
        return null;
    }

    private static InfoBadge CreateStatusBadge(bool signedIn)
    {
        var badge = new InfoBadge { Width = 12, Height = 12, MinWidth = 12, MinHeight = 12, Padding = new Thickness(0) };
        ApplyStatusBadge(badge, signedIn);
        return badge;
    }

    private static void ApplyStatusBadge(InfoBadge badge, bool signedIn)
    {
        badge.Style = (Style)Application.Current.Resources[signedIn ? "SuccessIconInfoBadgeStyle" : "CautionIconInfoBadgeStyle"];
        badge.IconSource = new FontIconSource
        {
            FontFamily = FluentIcons.FilledFontFamily,
            FontSize = 9,
            Glyph = signedIn ? FluentIcons.CheckmarkBadge : FluentIcons.ImportantBadge,
        };
    }

    /// <summary>
    /// Adds (<paramref name="profile"/> null) or edits a profile. A new profile, or one pointed at a different
    /// site, tenant or app, becomes current and connects.
    /// </summary>
    public async Task EditProfileAsync(ConnectionProfile? profile)
    {
        if (_profileDialogOpen) return;
        if (!await ConfirmLeaveAsync()) return;
        _profileDialogOpen = true;
        ConnectionProfile? toConnect = null;
        try
        {
            var dialog = new ProfileDialog(profile, App.Settings.Data.DefaultClientId) { XamlRoot = Content.XamlRoot, RequestedTheme = DialogTheme };
            if (await dialog.ShowAsync() != ContentDialogResult.Primary || dialog.Result is not { } saved) return;

            if (dialog.ConnectionChanged && profile != null) saved.AccountId = ""; // the old account may not apply any more
            App.Settings.SaveProfile(saved);
            if (profile == null || dialog.ConnectionChanged) toConnect = saved;
        }
        catch (Exception)
        {
            return; // another dialog is already open
        }
        finally
        {
            _profileDialogOpen = false;
        }

        UpdateProfileDisplay();
        if (toConnect != null)
        {
            App.Settings.SetCurrentProfile(toConnect);
            ClearTree();
            ShowTermStore();
            await ConnectAsync();
        }
    }

    private async Task SwitchProfileAsync(ConnectionProfile profile)
    {
        if (profile == CurrentProfile || !await ConfirmLeaveAsync()) return;
        App.Settings.SetCurrentProfile(profile);
        ClearTree();
        ShowTermStore();
        await ConnectAsync();
    }

    /// <summary>Signs out of the current profile and closes its term store.</summary>
    public async Task SignOutAsync()
    {
        if (CurrentProfile is not { } profile || !await ConfirmLeaveAsync()) return;
        await App.Auth.SignOutAsync(profile);
        App.Settings.SaveProfile(profile);
        ClearTree();
        UpdateProfileDisplay();
        ShowStatus($"Signed out of “{profile.EffectiveName}”.", InfoBarSeverity.Success);
    }

    private async Task ChangeAccountAsync()
    {
        if (CurrentProfile is not { } profile || !await ConfirmLeaveAsync()) return;
        try
        {
            await App.Auth.GetAccessTokenAsync(profile, allowInteractive: true, chooseAccount: true);
            App.Settings.SaveProfile(profile);
        }
        catch (Exception ex)
        {
            ShowError("Could not sign in", ex);
            return;
        }
        ClearTree();
        ShowTermStore();
        await ConnectAsync();
    }

    /// <summary>
    /// Settings › Reset, as in SharePoint Search Explorer: signs every profile out, forgets the sign-in window's
    /// session and restores all settings to their defaults (removing all profiles), optionally deletes the
    /// operations log, then starts over like a first run. Unsaved edits are offered for saving first; cancelling
    /// that cancels the reset. Nothing in SharePoint is changed.
    /// </summary>
    public async Task ResetToDefaultsAsync(bool deleteLog)
    {
        if (!await ConfirmLeaveAsync()) return;

        _loadCts?.Cancel();
        _detailsCts?.Cancel();
        foreach (var profile in App.Settings.Data.Profiles.ToList())
        {
            try
            {
                await App.Auth.SignOutAsync(profile);
            }
            catch (Exception)
            {
                // The settings reset matters more than clearing a cached token.
            }
        }

        App.Settings.ResetToDefaults();
        WebViewSignIn.ClearBrowsingData();
        if (deleteLog) OperationLog.DeleteAll();

        var data = App.Settings.Data;
        DisplayOptions.Current.UseColorfulIcons = data.UseColorfulIcons;
        ApplyTheme(data.Theme);
        ApplyAccentColor(data.AccentColor);
        ClearTree();
        GlobalSearchBox.Text = "";
        _statusTimer.Stop();
        StatusBar.IsOpen = false;
        SettingsView.LoadSettings();
        ExportView.LoadOptions();
        UpdateProfileDisplay();
        ShowTermStore();

        await EditProfileAsync(null);
    }

    private async Task RemoveProfileAsync(ConnectionProfile profile)
    {
        if (!await ConfirmLeaveAsync()) return;
        var confirm = new ContentDialog
        {
            XamlRoot = Content.XamlRoot,
            RequestedTheme = DialogTheme,
            Title = "Remove profile?",
            Content = $"Remove “{profile.EffectiveName}” and sign out of it? Nothing in SharePoint is changed.",
            PrimaryButtonText = "Remove",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Close,
        };
        if (await confirm.ShowAsync() != ContentDialogResult.Primary) return;

        bool wasCurrent = profile == CurrentProfile;
        await App.Auth.SignOutAsync(profile);
        App.Settings.RemoveProfile(profile);
        if (wasCurrent) ClearTree();
        UpdateProfileDisplay();
    }

    // ---- Term store tree -----------------------------------------------------------------------------------

    private async void Connect_Click(object sender, RoutedEventArgs e) => await RefreshAsync();

    private async void Refresh_Click(object sender, RoutedEventArgs e) => await RefreshAsync();

    /// <summary>Connects, or reloads the term store, after offering to save unsaved changes.</summary>
    private async Task RefreshAsync()
    {
        if (CurrentProfile == null)
        {
            await EditProfileAsync(null);
            return;
        }
        if (_busy || !await ConfirmLeaveAsync()) return;
        ShowTermStore();
        await ConnectAsync();
    }

    private async Task ConnectAsync()
    {
        if (CurrentProfile is not { } profile) return;

        _loadCts?.Cancel();
        _loadCts = new CancellationTokenSource();
        var ct = _loadCts.Token;

        SetBusy(true);
        StatusBar.IsOpen = false;
        try
        {
            var store = await App.TermStore.GetTermStoreAsync(profile, ct);
            if (ct.IsCancellationRequested) return;

            ClearTree();
            foreach (var group in store.Groups) Roots.Add(group);
            _storeLoaded = true;
            App.Settings.SaveProfile(profile); // the signed-in account may have changed
        }
        catch (Exception ex) when (!ct.IsCancellationRequested)
        {
            ShowError("Could not open the term store", ex);
        }
        finally
        {
            SetBusy(false);
            UpdateProfileDisplay();
        }
    }

    private async void Tree_Expanding(TreeView sender, TreeViewExpandingEventArgs args)
    {
        if (args.Item is not TermNode node || !node.HasUnrealizedChildren || node.IsLoading) return;
        if (CurrentProfile is not { } profile) return;

        try
        {
            await LoadChildrenAsync(profile, node);
        }
        catch (Exception ex)
        {
            args.Node.IsExpanded = false;
            ShowError($"Could not load the contents of “{node.Name}”", ex);
        }
    }

    private static async Task LoadChildrenAsync(ConnectionProfile profile, TermNode node)
    {
        node.IsLoading = true;
        try
        {
            // All the children are read (sorting needs them all), but the tree only shows a page at a time
            // (Settings › Options › Term tree), so opening a large term set stays quick.
            var children = await App.TermStore.GetChildrenAsync(profile, node);
            node.SetChildren(children, App.Settings.Data.TreePageSize);
            node.HasUnrealizedChildren = false;
        }
        finally
        {
            node.IsLoading = false;
        }
    }

    private async void Tree_ItemInvoked(TreeView sender, TreeViewItemInvokedEventArgs args)
    {
        if (args.InvokedItem is not TermNode node) return;
        if (node.IsMoreItem)
        {
            // "Show earlier" / "Show more": the adjacent page of children, keeping the selection on the item being edited.
            if (node.IsEarlierItem) node.Parent?.ShowEarlier();
            else node.Parent?.ShowMore();
            TrySelectInTree(_selected);
            return;
        }
        await SelectAsync(node);
    }

    private void Tree_KeyDown(object sender, KeyRoutedEventArgs e)
    {
        var node = Tree.SelectedItem as TermNode ?? _selected;
        if (node == null || node.IsMoreItem) return;
        switch (e.Key)
        {
            case VirtualKey.Delete:
                e.Handled = true;
                _ = DeleteAsync(node);
                break;
            case VirtualKey.F2:
                e.Handled = true;
                _ = RenameAsync(node);
                break;
        }
    }

    /// <summary>Shows a node in the editor, after offering to save unsaved changes. False if the user stayed.</summary>
    private async Task<bool> SelectAsync(TermNode node)
    {
        if (node == _selected && Editor.Original != null) return true;
        if (!await ConfirmLeaveAsync())
        {
            TrySelectInTree(_selected);
            return false;
        }
        TrySelectInTree(node);
        await ShowNodeAsync(node);
        return true;
    }

    /// <summary>Reads the node's details from SharePoint into the editor.</summary>
    private async Task ShowNodeAsync(TermNode node)
    {
        if (CurrentProfile is not { } profile) return;

        _selected = node;
        UpdateCommands();
        _detailsCts?.Cancel();
        var cts = _detailsCts = new CancellationTokenSource();
        Editor.ShowLoading(node);
        try
        {
            var details = await App.TermStore.GetDetailsAsync(profile, node, cts.Token);
            if (cts.IsCancellationRequested || _selected != node) return;
            SyncNodes(details, node);
            Editor.Show(node, details);
            UpdateEditorCommands(node);
        }
        catch (Exception ex) when (!cts.IsCancellationRequested)
        {
            if (_selected == node) Editor.Clear();
            ShowError($"Could not read “{node.Name}”", ex);
        }
    }

    /// <summary>
    /// Updates every loaded copy of the item (a reused term appears once per term set) with fresh values, and
    /// the reuse and pin flags of <paramref name="instance"/>, the copy the details were read from.
    /// </summary>
    private void SyncNodes(NodeDetails details, TermNode? instance = null)
    {
        foreach (var node in Roots.SelectMany(r => r.SelfAndDescendants()).Where(n => n.Id == details.Id && n.Kind == details.Kind))
        {
            node.Name = details.Name;
            node.Description = details.Description;
            if (node.Kind == TermNodeKind.Term) node.IsDeprecated = details.IsDeprecated;
        }
        if (instance is { Kind: TermNodeKind.Term })
        {
            instance.IsReused = details.IsReused;
            instance.IsPinned = details.IsPinned;
            instance.IsSourceTerm = details.IsSourceTerm;
            instance.ChildTermCount = details.TermsCount;
        }
    }

    private void ClearTree()
    {
        _detailsCts?.Cancel();
        Roots.Clear();
        _storeLoaded = false;
        _selected = null;
        Editor.Clear();
        UpdateCommands();
    }

    // ---- Create, rename, save, delete ----------------------------------------------------------------------

    private async void NewGroup_Click(object sender, RoutedEventArgs e) => await CreateAsync(TermNodeKind.Group, null);

    private void Tree_ContextRequested(UIElement sender, ContextRequestedEventArgs args)
    {
        var item = FindAncestor<TreeViewItem>(args.OriginalSource as DependencyObject);
        if (item == null || (Tree.ItemFromContainer(item) ?? item.DataContext) is not TermNode node) return;
        if (_busy || !_storeLoaded || node.IsMoreItem) return;
        args.Handled = true;

        var menu = new MenuFlyout();
        FillNodeMenu(menu, node);
        if (args.TryGetPosition(item, out var point))
            menu.ShowAt(item, new FlyoutShowOptions { Position = point });
        else
            menu.ShowAt(item);
    }

    /// <summary>
    /// One command for an item. <see cref="Label"/> is the menu text ("Move to…": the "…" says a dialog opens);
    /// <see cref="BarLabel"/> is the shorter command bar text ("Move"), where the full text is the tooltip.
    /// <see cref="GroupStart"/> puts a separator before it.
    /// </summary>
    private sealed record NodeCommand(
        string Label, string Glyph, bool Enabled, Func<Task> Run, bool GroupStart = false, string? BarLabel = null, bool IsDelete = false);

    /// <summary>The commands for one item, shared by the right-click menu and the editor's command bar.</summary>
    private List<NodeCommand> CommandsFor(TermNode node)
    {
        var commands = new List<NodeCommand>();
        switch (node.Kind)
        {
            case TermNodeKind.Group:
                commands.Add(new("New term set", FluentIcons.Add, CanCreateTermSetIn(node), () => CreateAsync(TermNodeKind.TermSet, node)));
                commands.Add(new("Rename", FluentIcons.Rename, CanEdit(node), () => RenameAsync(node)));
                break;
            case TermNodeKind.TermSet:
                commands.Add(new("New term", FluentIcons.Add, CanCreateTermUnder(node), () => CreateAsync(TermNodeKind.Term, node)));
                commands.Add(new("Rename", FluentIcons.Rename, CanEdit(node), () => RenameAsync(node)));
                commands.Add(new("Move to group…", FluentIcons.ArrowMove, CanMove(node), () => MoveAsync(node), GroupStart: true, BarLabel: "Move"));
                break;
            default:
                commands.Add(new("New child term", FluentIcons.Add, CanCreateTermUnder(node), () => CreateAsync(TermNodeKind.Term, node)));
                commands.Add(new("Rename", FluentIcons.Rename, CanEdit(node), () => RenameAsync(node)));
                commands.Add(new("Move to…", FluentIcons.ArrowMove, CanMove(node), () => MoveAsync(node), GroupStart: true, BarLabel: "Move"));
                commands.Add(new("Merge into…", FluentIcons.Merge, CanMerge(node), () => MergeAsync(node), BarLabel: "Merge"));
                commands.Add(node.IsDeprecated
                    ? new("Un-deprecate", FluentIcons.Tag, CanDeprecate(node), () => SetDeprecatedAsync(node, false))
                    : new("Deprecate", FluentIcons.TagOff, CanDeprecate(node), () => SetDeprecatedAsync(node, true)));
                commands.Add(new("Copy…", FluentIcons.Copy, CanCopy(node), () => CopyAsync(node), BarLabel: "Copy"));
                commands.Add(new("Reuse in…", FluentIcons.Link, true, () => ReuseAsync(node, pin: false), BarLabel: "Reuse"));
                commands.Add(new("Pin to…", FluentIcons.Pin, true, () => ReuseAsync(node, pin: true), BarLabel: "Pin"));
                break;
        }
        commands.Add(new("Delete…", FluentIcons.Delete, CanDelete(node), () => DeleteAsync(node), GroupStart: true, BarLabel: "Delete", IsDelete: true));
        return commands;
    }

    /// <summary>
    /// The details pane's commands: the item's commands (without Delete when Settings › Options hides it there),
    /// then Close.
    /// </summary>
    private List<ICommandBarElement> EditorBarCommands(TermNode node, bool enabled)
    {
        var commands = CommandsFor(node).Where(c => !c.IsDelete || App.Settings.Data.ShowDeleteInDetailsPane).ToList();
        commands.Add(new("Close the details pane", FluentIcons.Dismiss, true, CloseEditorAsync, GroupStart: true, BarLabel: "Close"));
        return BarCommands(commands, enabled);
    }

    /// <summary>Clears the details pane and the tree's selection, after offering to save unsaved changes.</summary>
    private async Task CloseEditorAsync()
    {
        if (_busy || !await ConfirmLeaveAsync()) return;
        _detailsCts?.Cancel();
        _selected = null;
        Editor.Clear();
        TrySelectInTree(null);
        UpdateCommands();
    }

    /// <summary>Rebuilds the details pane's command bar, e.g. after the Delete button setting changes.</summary>
    public void RefreshEditorCommands()
    {
        if (Editor.Node == null) Editor.SetPlaceholderCommands(PlaceholderCommands);
        else if (Editor.Original != null) UpdateEditorCommands(Editor.Node);
    }

    /// <summary>The right-click menu for an item.</summary>
    private void FillNodeMenu(MenuFlyout menu, TermNode node)
    {
        foreach (var command in CommandsFor(node))
        {
            if (command.GroupStart && menu.Items.Count > 0) menu.Items.Add(new MenuFlyoutSeparator());
            var item = new MenuFlyoutItem { Text = command.Label, Icon = FluentIcons.Create(command.Glyph), IsEnabled = command.Enabled };
            item.Click += async (_, _) => await command.Run();
            menu.Items.Add(item);
        }
    }

    /// <summary>
    /// Fills the editor's command bar with the item's commands. They're all primary commands; when the panel is
    /// too narrow for them, the bar moves the last ones into its "…" menu by itself.
    /// </summary>
    private void UpdateEditorCommands(TermNode node) => Editor.SetItemCommands(EditorBarCommands(node, enabled: true));

    /// <summary>
    /// The command bar shown, all disabled, while no item is selected: a term's commands (the fullest set), so
    /// people can see what they'll be able to do once they pick an item in the tree.
    /// </summary>
    private List<ICommandBarElement> PlaceholderCommands() =>
        EditorBarCommands(new TermNode { Kind = TermNodeKind.Term }, enabled: false);

    /// <summary>Command bar buttons for <paramref name="nodeCommands"/>; when not <paramref name="enabled"/>, all disabled and inert.</summary>
    private static List<ICommandBarElement> BarCommands(IEnumerable<NodeCommand> nodeCommands, bool enabled)
    {
        var commands = new List<ICommandBarElement>();
        foreach (var command in nodeCommands)
        {
            if (command.GroupStart && commands.Count > 0) commands.Add(new AppBarSeparator());
            var button = new AppBarButton { Label = command.BarLabel ?? command.Label, Icon = FluentIcons.Create(command.Glyph), IsEnabled = enabled && command.Enabled };
            ToolTipService.SetToolTip(button, command.Label);
            if (enabled) button.Click += async (_, _) => await command.Run();
            commands.Add(button);
        }
        return commands;
    }


    private async Task RenameAsync(TermNode node)
    {
        if (CanEdit(node) && await SelectAsync(node)) Editor.FocusName();
    }

    private async Task CreateAsync(TermNodeKind kind, TermNode? parent)
    {
        if (CurrentProfile is not { } profile || _busy || !_storeLoaded) return;
        if (!await ConfirmLeaveAsync()) return;

        // Load what's already there first, to check for duplicate names and to show the new item among them.
        if (parent is { HasUnrealizedChildren: true })
        {
            SetBusy(true);
            try
            {
                await LoadChildrenAsync(profile, parent);
            }
            catch (Exception ex)
            {
                ShowError($"Could not load the contents of “{parent.Name}”", ex);
                return;
            }
            finally
            {
                SetBusy(false);
            }
        }

        var siblings = parent?.LoadedChildren.ToList() ?? (IReadOnlyCollection<TermNode>)Roots;
        var dialog = new NewItemDialog(kind, parent, siblings) { XamlRoot = Content.XamlRoot, RequestedTheme = DialogTheme };
        if (await dialog.ShowAsync() != ContentDialogResult.Primary) return;

        SetBusy(true);
        TermNode node;
        try
        {
            node = await App.TermStore.CreateAsync(profile, parent, kind, dialog.ItemName, dialog.ItemDescription);
            OperationLog.Write(profile, "create", "succeeded",
                new { item = node.ToRef(), parent = parent?.ToRef(), description = dialog.ItemDescription });
        }
        catch (Exception ex)
        {
            // No ID yet: log what was asked for.
            var path = parent == null ? dialog.ItemName : $"{parent.Path} › {dialog.ItemName}";
            OperationLog.Write(profile, "create", "failed",
                new { item = new { Kind = kind, Name = dialog.ItemName, Path = path }, parent = parent?.ToRef(), description = dialog.ItemDescription }, ex);
            ShowError($"Could not create the {TermNode.Noun(kind)} “{dialog.ItemName}”", ex);
            return;
        }
        finally
        {
            SetBusy(false);
        }

        InsertSorted(node);
        if (parent != null)
        {
            parent.HasUnrealizedChildren = false;
            if (parent.Kind == TermNodeKind.Term) parent.ChildTermCount++;
            parent.IsExpanded = true;
        }
        ShowStatus($"Created the {TermNode.Noun(kind)} “{node.Name}”.", InfoBarSeverity.Success);
        TrySelectInTree(node);
        await ShowNodeAsync(node);
    }

    /// <summary>Saves the editor's changes. True when there was nothing to save or the save succeeded.</summary>
    private async Task<bool> SaveAsync()
    {
        if (Editor.Node is not { } node || Editor.Original is not { } original || CurrentProfile is not { } profile) return false;
        if (_busy) return false;

        var edited = Editor.GetEdited(out var error);
        if (edited == null)
        {
            Editor.ShowValidationError(error ?? "Check the values you entered.");
            return false;
        }
        if (edited.SameEditableValues(original)) return true;

        var siblings = node.Parent?.LoadedChildren ?? Roots;
        if (!string.Equals(edited.Name, original.Name, StringComparison.OrdinalIgnoreCase)
            && siblings.Any(s => s != node && string.Equals(s.Name, edited.Name, StringComparison.OrdinalIgnoreCase)))
        {
            Editor.ShowValidationError($"There's already an item named “{edited.Name}” at this level.");
            return false;
        }

        // Logged with the item as it was before the save, and each change's old and new value.
        var logData = new { item = node.ToRef(), language = original.Language, changes = original.ChangesTo(edited) };

        SetBusy(true);
        Editor.SetBusy(true);
        try
        {
            await App.TermStore.UpdateAsync(profile, node, original, edited);
            OperationLog.Write(profile, "update", "succeeded", logData);
        }
        catch (Exception ex)
        {
            OperationLog.Write(profile, "update", "failed", logData, ex);
            ShowError($"Could not save “{original.Name}”", ex);
            return false;
        }
        finally
        {
            Editor.SetBusy(false);
            SetBusy(false);
        }

        // Show what SharePoint now holds (names may be normalized; the modified date is needed for the next save).
        try
        {
            var fresh = await App.TermStore.GetDetailsAsync(profile, node);
            SyncNodes(fresh, node);
            if (_selected == node)
            {
                Editor.Show(node, fresh);
                UpdateEditorCommands(node);
            }
            ShowStatus($"Saved “{fresh.Name}”.", InfoBarSeverity.Success);
        }
        catch (Exception)
        {
            node.Name = edited.Name;
            if (_selected == node) Editor.Clear();
            ShowStatus($"Saved “{edited.Name}”, but couldn't read it back. Select it again to see the latest values.", InfoBarSeverity.Warning);
        }
        return true;
    }

    private async Task DeleteAsync(TermNode node)
    {
        if (CurrentProfile is not { } profile || _busy || !CanDelete(node)) return;
        // Unsaved edits to the item being deleted don't matter; edits to another item do.
        if (node != Editor.Node && !await ConfirmLeaveAsync()) return;

        if (node.Kind == TermNodeKind.Group && !node.HasUnrealizedChildren && node.LoadedChildren.Any())
        {
            ShowStatus($"“{node.Name}” still contains term sets. Delete or move them first; SharePoint only deletes empty term groups.", InfoBarSeverity.Warning);
            return;
        }

        string question = node.Kind switch
        {
            TermNodeKind.Group => $"Delete the term group “{node.Name}”?",
            TermNodeKind.TermSet => $"Delete the term set “{node.Name}” and every term in it?",
            _ when node.ChildTermCount > 0 => $"Delete the term “{node.Name}” and all {node.ChildTermCount} term(s) below it, including their children?",
            _ => $"Delete the term “{node.Name}”?",
        };
        if (node.Kind == TermNodeKind.Term && node.IsReused && !node.IsSourceTerm)
            question += " Only this reused copy in “" + node.TermSet?.Name + "” is deleted.";
        if (node.Kind != TermNodeKind.Group)
            question += "\n\nContent already tagged with these terms isn't updated.";
        question += "\n\nThis can't be undone.";

        var confirm = new ContentDialog
        {
            XamlRoot = Content.XamlRoot,
            RequestedTheme = DialogTheme,
            Title = $"Delete {TermNode.Noun(node.Kind)}?",
            Content = new TextBlock { Text = question, TextWrapping = TextWrapping.Wrap },
            PrimaryButtonText = "Delete",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Close,
        };
        if (await confirm.ShowAsync() != ContentDialogResult.Primary) return;

        SetBusy(true);
        var logData = new { item = node.ToRef(), snapshot = await TrySnapshotAsync(profile, node) };
        OperationLog.Write(profile, "delete", "started", logData);
        try
        {
            await App.TermStore.DeleteAsync(profile, node);
            OperationLog.Write(profile, "delete", "succeeded", new { item = node.ToRef() });
        }
        catch (Exception ex)
        {
            OperationLog.Write(profile, "delete", "failed", new { item = node.ToRef() }, ex);
            ShowError($"Could not delete “{node.Name}”", ex);
            return;
        }
        finally
        {
            SetBusy(false);
        }

        RemoveNode(node);
        if (_selected != null && node.SelfAndDescendants().Contains(_selected))
        {
            _detailsCts?.Cancel();
            _selected = null;
            Editor.Clear();
        }
        UpdateCommands();
        ShowStatus($"Deleted the {TermNode.Noun(node.Kind)} “{node.Name}”.", InfoBarSeverity.Success);
    }

    /// <summary>
    /// If the editor has unsaved changes, asks whether to save them. True to carry on (saved or discarded),
    /// false to stay.
    /// </summary>
    private async Task<bool> ConfirmLeaveAsync()
    {
        if (!Editor.IsDirty || Editor.Node is not { } node) return true;

        var dialog = new ContentDialog
        {
            XamlRoot = Content.XamlRoot,
            RequestedTheme = DialogTheme,
            Title = "Save changes?",
            Content = $"“{node.Name}” has changes that haven't been saved to SharePoint.",
            PrimaryButtonText = "Save",
            SecondaryButtonText = "Don't save",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Primary,
        };
        ContentDialogResult result;
        try
        {
            result = await dialog.ShowAsync();
        }
        catch (Exception)
        {
            return false; // another dialog is already open
        }

        return result switch
        {
            ContentDialogResult.Primary => await SaveAsync() && !Editor.IsDirty,
            ContentDialogResult.Secondary => true,
            _ => false,
        };
    }

    private async void AppWindow_Closing(Microsoft.UI.Windowing.AppWindow sender, Microsoft.UI.Windowing.AppWindowClosingEventArgs args)
    {
        if (_closeConfirmed || !Editor.IsDirty) return;
        args.Cancel = true;
        if (await ConfirmLeaveAsync())
        {
            _closeConfirmed = true;
            Close();
        }
    }

    // ---- Deprecate, merge, move, copy, reuse, pin ----------------------------------------------------------

    private async Task SetDeprecatedAsync(TermNode node, bool deprecate)
    {
        if (CurrentProfile is not { } profile || _busy || !CanDeprecate(node)) return;
        if (!await ConfirmLeaveAsync()) return;

        string verb = deprecate ? "deprecate" : "un-deprecate";
        var item = node.ToRef();
        SetBusy(true);
        try
        {
            node.IsDeprecated = await App.TermStore.SetDeprecatedAsync(profile, item, deprecate);
            OperationLog.Write(profile, verb, "succeeded", new { item });
        }
        catch (Exception ex)
        {
            OperationLog.Write(profile, verb, "failed", new { item }, ex);
            ShowError($"Could not {verb} “{node.Name}”", ex);
            return;
        }
        finally
        {
            SetBusy(false);
        }

        // Child terms may have changed with it.
        if (!node.HasUnrealizedChildren && node.LoadedChildren.Any()) await ReloadChildrenAsync(profile, node, expand: false);
        await ShowIfPresentAsync(_selected?.ToRef());
        ShowStatus(deprecate
            ? $"Deprecated “{node.Name}”. It stays in the term store but can't be used for new tags."
            : $"“{node.Name}” can be used for tagging again.", InfoBarSeverity.Success);
    }

    private async Task MergeAsync(TermNode node)
    {
        if (CurrentProfile is not { } profile || _busy || !CanMerge(node)) return;
        if (!await ConfirmLeaveAsync()) return;
        if (node.ChildTermCount > 0)
        {
            ShowStatus($"“{node.Name}” has child terms. Move or merge them first, so it's clear where they end up.", InfoBarSeverity.Warning);
            return;
        }

        var source = node.ToRef();
        var picker = new TermPickerDialog(profile, new PickerOptions
        {
            Title = $"Merge “{node.Name}” into…",
            Intro = $"Choose the term to keep. “{node.Name}” is merged into it and then removed.",
            PrimaryButtonText = "Next",
            Allowed = new HashSet<TermNodeKind> { TermNodeKind.Term },
            ScopeTermSet = node.TermSet?.ToRef(),
            Validate = t =>
                t.Id == source.Id ? "That's the term being merged. Pick the term to merge it into."
                : t.IsPinnedCopy ? "That's a pinned copy. Pick the term in its source term set."
                : IsAtOrBelow(t, source) ? "That term is below the one being merged."
                : null,
        }) { XamlRoot = Content.XamlRoot, RequestedTheme = DialogTheme };
        if (await picker.ShowAsync() != ContentDialogResult.Primary || picker.Result is not { } target) return;

        // Read the source again for the preview and the log: its labels are what the target gains.
        SetBusy(true);
        NodeDetails? details;
        try
        {
            details = await App.TermStore.GetDetailsAsync(profile, node);
        }
        catch (Exception ex)
        {
            ShowError($"Could not read “{node.Name}”", ex);
            return;
        }
        finally
        {
            SetBusy(false);
        }
        if (details.TermsCount > 0)
        {
            ShowStatus($"“{node.Name}” has child terms. Move or merge them first.", InfoBarSeverity.Warning);
            return;
        }

        var labels = new[] { details.Name }.Concat(details.Synonyms).Select(l => $"“{l}”");
        var lines = new List<string>
        {
            $"From:  {source.Path}",
            $"Into:  {target.Path}",
            "",
            $"• {string.Join(", ", labels)} become labels (synonyms) of “{target.Name}”.",
            $"• “{node.Name}” is removed. Its ID is kept on “{target.Name}” as a merged term ID, so content tagged with it can be resolved to “{target.Name}”.",
        };
        if (details.Description.Length > 0 || details.CustomProperties.Count > 0 || details.LocalCustomProperties.Count > 0)
            lines.Add($"• Its description and custom properties may not be kept. A copy is saved to the operations log.");
        if (details.IsReused)
            lines.Add($"• “{node.Name}” is reused in other term sets. SharePoint may refuse to merge it.");
        lines.Add("");
        lines.Add("This can't be undone.");

        var confirm = new ContentDialog
        {
            XamlRoot = Content.XamlRoot,
            RequestedTheme = DialogTheme,
            Title = $"Merge “{node.Name}” into “{target.Name}”?",
            Content = new TextBlock { Text = string.Join("\n", lines), TextWrapping = TextWrapping.Wrap },
            PrimaryButtonText = "Merge",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Close,
        };
        if (await confirm.ShowAsync() != ContentDialogResult.Primary) return;

        SetBusy(true);
        OperationLog.Write(profile, "merge", "started", new { source, target, snapshot = details });
        try
        {
            await App.TermStore.MergeAsync(profile, source, target);
            OperationLog.Write(profile, "merge", "succeeded", new { source, target });
        }
        catch (Exception ex)
        {
            OperationLog.Write(profile, "merge", "failed", new { source, target }, ex);
            ShowError($"Could not merge “{node.Name}” into “{target.Name}”", ex);
            return;
        }
        finally
        {
            SetBusy(false);
        }

        RemoveNode(node);
        await ShowIfPresentAsync(target);
        ShowStatus($"Merged “{node.Name}” into “{target.Name}”." + (node.IsReused ? " Refresh to update its other term sets." : ""),
            InfoBarSeverity.Success);
    }

    private async Task MoveAsync(TermNode node)
    {
        if (CurrentProfile is not { } profile || _busy || !CanMove(node)) return;
        if (!await ConfirmLeaveAsync()) return;

        var item = node.ToRef();
        var currentParent = node.Parent?.ToRef();
        bool isTermSet = node.Kind == TermNodeKind.TermSet;

        var picker = new TermPickerDialog(profile, isTermSet
            ? new PickerOptions
            {
                Title = $"Move “{node.Name}” to…",
                Intro = $"Choose the term group to move the term set “{node.Name}” and all its terms to.",
                PrimaryButtonText = "Move",
                Allowed = new HashSet<TermNodeKind> { TermNodeKind.Group },
                Deepest = TermNodeKind.Group,
                Validate = g =>
                    g.IsSystemGroup ? "Term sets can't be moved into a system group."
                    : g.Id == currentParent?.Id ? "It's already in that group."
                    : null,
            }
            : new PickerOptions
            {
                Title = $"Move “{node.Name}” to…",
                Intro = $"Choose the term set, or the term to put it under. “{node.Name}” moves with everything below it.",
                PrimaryButtonText = "Move",
                Allowed = new HashSet<TermNodeKind> { TermNodeKind.TermSet, TermNodeKind.Term },
                ScopeTermSet = node.TermSet?.ToRef(),
                Validate = t =>
                    t.Kind == TermNodeKind.Term && t.Id == item.Id ? "That's the term being moved."
                    : IsAtOrBelow(t, item) ? "A term can't be moved below itself."
                    : currentParent != null && SameItem(t, currentParent) ? "It's already there."
                    : t.IsPinnedCopy ? "Terms can't be added under a pinned copy."
                    : null,
            }) { XamlRoot = Content.XamlRoot, RequestedTheme = DialogTheme };
        if (await picker.ShowAsync() != ContentDialogResult.Primary || picker.Result is not { } target) return;

        SetBusy(true);
        var logData = new { item, from = currentParent, to = target };
        try
        {
            if (isTermSet) await App.TermStore.MoveTermSetAsync(profile, item, target);
            else await App.TermStore.MoveTermAsync(profile, item, target);
            OperationLog.Write(profile, "move", "succeeded", logData);
        }
        catch (Exception ex)
        {
            OperationLog.Write(profile, "move", "failed", logData, ex);
            ShowError($"Could not move “{node.Name}”", ex);
            return;
        }
        finally
        {
            SetBusy(false);
        }

        RemoveNode(node);
        await ReloadChildrenAsync(profile, target, expand: true);
        var newTermSetId = target.Kind == TermNodeKind.TermSet ? target.Id : target.TermSetId;
        await ShowIfPresentAsync(isTermSet ? item : item with { TermSetId = newTermSetId });
        ShowStatus($"Moved “{node.Name}” to {target.Path}.", InfoBarSeverity.Success);
    }

    private async Task CopyAsync(TermNode node)
    {
        if (CurrentProfile is not { } profile || _busy || !CanCopy(node) || node.Parent is not { } parent) return;
        if (!await ConfirmLeaveAsync()) return;

        var withChildren = new CheckBox
        {
            Content = node.ChildTermCount > 0 ? $"Include its child terms ({node.ChildTermCount} directly below it)" : "Include child terms (it has none)",
            IsChecked = node.ChildTermCount > 0,
            IsEnabled = node.ChildTermCount > 0,
        };
        var dialog = new ContentDialog
        {
            XamlRoot = Content.XamlRoot,
            RequestedTheme = DialogTheme,
            Title = $"Copy “{node.Name}”?",
            Content = new StackPanel
            {
                Spacing = 12,
                Children =
                {
                    new TextBlock
                    {
                        Text = "SharePoint creates a new, separate term next to the original and names it. You can rename or move the copy afterwards. Tags on content stay with the original.",
                        TextWrapping = TextWrapping.Wrap,
                    },
                    withChildren,
                },
            },
            PrimaryButtonText = "Copy",
            CloseButtonText = "Cancel",
            DefaultButton = ContentDialogButton.Primary,
        };
        if (await dialog.ShowAsync() != ContentDialogResult.Primary) return;

        var item = node.ToRef();
        bool includeChildren = withChildren.IsChecked == true;
        SetBusy(true);
        TermNode copy;
        try
        {
            copy = await App.TermStore.CopyTermAsync(profile, item, parent, includeChildren);
            OperationLog.Write(profile, "copy", "succeeded", new { item, includeChildren, copy = copy.ToRef() });
        }
        catch (Exception ex)
        {
            OperationLog.Write(profile, "copy", "failed", new { item, includeChildren }, ex);
            ShowError($"Could not copy “{node.Name}”", ex);
            return;
        }
        finally
        {
            SetBusy(false);
        }

        InsertSorted(copy);
        if (parent.Kind == TermNodeKind.Term) parent.ChildTermCount++;
        TrySelectInTree(copy);
        await ShowNodeAsync(copy);
        ShowStatus($"Copied “{node.Name}” as “{copy.Name}”.", InfoBarSeverity.Success);
    }

    private async Task ReuseAsync(TermNode node, bool pin)
    {
        if (CurrentProfile is not { } profile || _busy || node.Kind != TermNodeKind.Term) return;
        if (!await ConfirmLeaveAsync()) return;

        var source = node.ToRef();
        var picker = new TermPickerDialog(profile, new PickerOptions
        {
            Title = pin ? $"Pin “{node.Name}” to…" : $"Reuse “{node.Name}” in…",
            Intro = pin
                ? $"Choose the term set, or the term to put it under. A pinned copy of “{node.Name}” and its child terms appears there; it can only be changed in “{node.TermSet?.Name}”, and changes there show up automatically."
                : $"Choose the term set, or the term to put it under. “{node.Name}” appears there as the same term: its labels, description and shared custom properties stay the same everywhere.",
            PrimaryButtonText = pin ? "Pin" : "Next",
            Allowed = new HashSet<TermNodeKind> { TermNodeKind.TermSet, TermNodeKind.Term },
            Validate = t =>
                (t.Kind == TermNodeKind.TermSet ? t.Id : t.TermSetId) == source.TermSetId
                    ? "A term can only appear once in a term set. Pick a different term set."
                : t.IsPinnedCopy ? "Terms can't be added under a pinned copy."
                : null,
        }) { XamlRoot = Content.XamlRoot, RequestedTheme = DialogTheme };
        if (await picker.ShowAsync() != ContentDialogResult.Primary || picker.Result is not { } target) return;

        bool includeChildren = pin;
        if (!pin && node.ChildTermCount > 0)
        {
            var withChildren = new CheckBox { Content = "Also reuse its child terms", IsChecked = true };
            var dialog = new ContentDialog
            {
                XamlRoot = Content.XamlRoot,
                RequestedTheme = DialogTheme,
                Title = $"Reuse “{node.Name}” in {target.Name}?",
                Content = new StackPanel
                {
                    Spacing = 12,
                    Children =
                    {
                        new TextBlock { Text = $"Into: {target.Path}", TextWrapping = TextWrapping.Wrap },
                        withChildren,
                    },
                },
                PrimaryButtonText = "Reuse",
                CloseButtonText = "Cancel",
                DefaultButton = ContentDialogButton.Primary,
            };
            if (await dialog.ShowAsync() != ContentDialogResult.Primary) return;
            includeChildren = withChildren.IsChecked == true;
        }

        string operation = pin ? "pin" : "reuse";
        var logData = new { source, target, includeChildren };
        SetBusy(true);
        try
        {
            await App.TermStore.ReuseAsync(profile, source, target, pin, includeChildren);
            OperationLog.Write(profile, operation, "succeeded", logData);
        }
        catch (Exception ex)
        {
            OperationLog.Write(profile, operation, "failed", logData, ex);
            ShowError($"Could not {operation} “{node.Name}”", ex);
            return;
        }
        finally
        {
            SetBusy(false);
        }

        if (!node.IsReused)
        {
            node.IsReused = true;
            node.IsSourceTerm = true;
        }
        await ReloadChildrenAsync(profile, target, expand: true);
        await ShowIfPresentAsync(_selected?.ToRef());
        ShowStatus(pin ? $"Pinned “{node.Name}” to {target.Path}." : $"Reused “{node.Name}” in {target.Path}.", InfoBarSeverity.Success);
    }

    /// <summary>The item's details for the operations log, or null if they can't be read.</summary>
    private static async Task<NodeDetails?> TrySnapshotAsync(ConnectionProfile profile, TermNode node)
    {
        try
        {
            return await App.TermStore.GetDetailsAsync(profile, node);
        }
        catch (Exception)
        {
            return null;
        }
    }

    // ---- Keeping the tree in step --------------------------------------------------------------------------

    private List<TermNode> AllNodes() => Roots.SelectMany(r => r.SelfAndDescendants()).ToList();

    private void RemoveNode(TermNode node)
    {
        if (node.Parent is { } parent) parent.RemoveChild(node);
        else Roots.Remove(node);
        if (node.Parent is { Kind: TermNodeKind.Term } parentTerm)
            parentTerm.ChildTermCount = Math.Max(0, parentTerm.ChildTermCount - 1);
    }

    /// <summary>Re-reads the children of every loaded copy of <paramref name="container"/> whose children were loaded.</summary>
    private async Task ReloadChildrenAsync(ConnectionProfile profile, ItemRef container, bool expand)
    {
        foreach (var node in AllNodes().Where(n => n.Matches(container)))
            await ReloadChildrenAsync(profile, node, expand);
    }

    private async Task ReloadChildrenAsync(ConnectionProfile profile, TermNode node, bool expand)
    {
        if (node.HasUnrealizedChildren) return; // read fresh when expanded
        try
        {
            await LoadChildrenAsync(profile, node);
            if (node.Kind == TermNodeKind.Term) node.ChildTermCount = node.LoadedChildren.Count();
            if (expand && node.Children.Count > 0) node.IsExpanded = true;
        }
        catch (Exception)
        {
            node.ClearChildren();
            node.HasUnrealizedChildren = true; // try again on next expand
        }
    }

    /// <summary>Shows <paramref name="item"/> in the editor if it's in the loaded tree; otherwise clears the editor.</summary>
    private async Task ShowIfPresentAsync(ItemRef? item)
    {
        var match = item == null ? null : AllNodes().FirstOrDefault(n => n.Matches(item));
        if (match == null)
        {
            _detailsCts?.Cancel();
            _selected = null;
            Editor.Clear();
            UpdateCommands();
            return;
        }
        TrySelectInTree(match);
        await ShowNodeAsync(match);
    }

    /// <summary>True when <paramref name="candidate"/> is <paramref name="term"/> or below it, judged by path.</summary>
    private static bool IsAtOrBelow(ItemRef candidate, ItemRef term) =>
        candidate.Path == term.Path || candidate.Path.StartsWith(term.Path + " › ", StringComparison.Ordinal);

    private static bool SameItem(ItemRef a, ItemRef b) =>
        a.Kind == b.Kind && a.Id == b.Id && (a.Kind != TermNodeKind.Term || a.TermSetId == b.TermSetId);

    // ---- Rules for what can be changed ---------------------------------------------------------------------

    private static TermNode? GroupOf(TermNode? node)
    {
        while (node?.Parent != null) node = node.Parent;
        return node;
    }

    private static bool InSystemGroup(TermNode node) => GroupOf(node)?.IsSystemGroup == true;

    private static bool CanCreateTermSetIn(TermNode? group) => group is { Kind: TermNodeKind.Group, IsSystemGroup: false };

    private static bool CanCreateTermUnder(TermNode? node) =>
        node is { Kind: TermNodeKind.TermSet } || node is { Kind: TermNodeKind.Term, IsPinnedCopy: false };

    private static bool CanEdit(TermNode node) => node.Kind switch
    {
        TermNodeKind.Group => !node.IsSystemGroup,
        TermNodeKind.Term => !node.IsPinnedCopy,
        _ => true,
    };

    /// <summary>System groups and their term sets (Keywords, Hashtags, Orphaned Terms, People…) are SharePoint's own.</summary>
    private static bool CanDelete(TermNode? node) => node switch
    {
        null => false,
        { Kind: TermNodeKind.Group } => !node.IsSystemGroup,
        { Kind: TermNodeKind.TermSet } => !InSystemGroup(node),
        _ => true,
    };

    /// <summary>Pinned copies can only be changed at their source.</summary>
    private static bool CanDeprecate(TermNode node) => node is { Kind: TermNodeKind.Term, IsPinnedCopy: false };

    private static bool CanMerge(TermNode node) => node is { Kind: TermNodeKind.Term, IsPinnedCopy: false };

    private static bool CanCopy(TermNode node) => node is { Kind: TermNodeKind.Term, IsPinnedCopy: false, Parent: not null };

    private static bool CanMove(TermNode node) => node.Kind switch
    {
        TermNodeKind.Term => !node.IsPinnedCopy,
        TermNodeKind.TermSet => !InSystemGroup(node),
        _ => false,
    };

    private void UpdateCommands()
    {
        bool ready = _storeLoaded && !_busy;
        NewGroupButton.IsEnabled = ready;
        Editor.SetCommandsEnabled(!_busy);
        RefreshButton.IsEnabled = ready;

        // Until a term store is loaded, the tree card fills the window, explains what to do and offers to
        // connect; the profile's name is shown above the message, so the message doesn't repeat it.
        var profile = CurrentProfile;
        UpdatePaneLayout(_storeLoaded);
        EmptyState.Visibility = _storeLoaded ? Visibility.Collapsed : Visibility.Visible;
        Tree.Visibility = _storeLoaded ? Visibility.Visible : Visibility.Collapsed;
        EmptyStateText.Text = profile == null
            ? "Add a SharePoint site profile to open its term store."
            : _busy ? "Opening the term store…"
            : "Connect to open this site's term store. Switch sites from the profile button at the bottom left.";
        ConnectButtonText.Text = profile == null ? "Add profile" : "Connect";
        ConnectButton.IsEnabled = !_busy;
        ConnectButton.Visibility = _busy ? Visibility.Collapsed : Visibility.Visible;
        // The header's busy ring is hidden until a store is loaded, so connecting shows one here instead.
        EmptyStateRing.IsActive = _busy && !_storeLoaded;
        EmptyStateRing.Visibility = EmptyStateRing.IsActive ? Visibility.Visible : Visibility.Collapsed;

        GlobalSearchBox.IsEnabled = GlobalSearchButton.IsEnabled = _storeLoaded;
        GlobalSearchBox.PlaceholderText = _storeLoaded
            ? "Search terms (label or synonym starts with), term sets and groups, or paste an ID"
            : "Connect to search terms, term sets and groups";

        ExportView.UpdateState(_storeLoaded, Roots);
    }

    // ---- Helpers -------------------------------------------------------------------------------------------

    /// <summary>Inserts a new node among its siblings in the order the tree uses.</summary>
    /// <remarks>
    /// Under a parent, the position is among all its loaded children; if that's in the part the tree doesn't
    /// show, it shows the page around the new item so it's visible.
    /// </remarks>
    private void InsertSorted(TermNode node)
    {
        var list = node.Parent?.LoadedChildren.ToList() ?? Roots.ToList();
        int index;
        if (node.Kind == TermNodeKind.Term && node.Parent?.HasCustomSortOrder == true)
        {
            index = list.Count; // custom-ordered: new terms go last, as in SharePoint
        }
        else
        {
            var compare = StringComparer.CurrentCultureIgnoreCase;
            index = 0;
            while (index < list.Count
                   && (list[index].IsSystemGroup == node.IsSystemGroup
                       ? compare.Compare(list[index].Name, node.Name) <= 0
                       : !list[index].IsSystemGroup))
                index++;
        }

        if (node.Parent is { } parent) parent.InsertChild(index, node);
        else Roots.Insert(index, node);
    }

    private void TrySelectInTree(TermNode? node)
    {
        try
        {
            Tree.SelectedItem = node;
        }
        catch (Exception)
        {
            // The item's container isn't realized yet; the editor still shows it.
        }
    }

    private static T? FindAncestor<T>(DependencyObject? element) where T : DependencyObject
    {
        while (element != null && element is not T) element = VisualTreeHelper.GetParent(element);
        return element as T;
    }

    // ---- Status --------------------------------------------------------------------------------------------

    private void SetBusy(bool busy)
    {
        _busy = busy;
        BusyRing.IsActive = busy;
        UpdateCommands();
    }

    /// <summary>Shows a message in the status bar. Success messages close themselves after the time set in Settings.</summary>
    public void ShowStatus(string message, InfoBarSeverity severity)
    {
        _statusTimer.Stop();
        StatusBar.Title = "";
        StatusBar.Message = message;
        StatusBar.Severity = severity;
        StatusBar.IsOpen = true;

        int seconds = App.Settings.Data.NotificationTimeoutSeconds;
        if (severity == InfoBarSeverity.Success && seconds > 0)
        {
            _statusTimer.Interval = TimeSpan.FromSeconds(seconds);
            _statusTimer.Start();
        }
    }


    private void ShowError(string title, Exception ex)
    {
        var message = ex switch
        {
            SignInCanceledException => "Sign-in was canceled.",
            SignInRequiredException => ex.Message,
            ServerUnauthorizedAccessException =>
                "Access denied. Check that your account can open this site and has permissions in the term store.",
            _ => ex.Message,
        };
        _statusTimer.Stop();
        StatusBar.Title = title;
        StatusBar.Message = message;
        StatusBar.Severity = InfoBarSeverity.Error;
        StatusBar.IsOpen = true;
    }

    // ---- Global search -------------------------------------------------------------------------------------

    private void GlobalSearchBox_TextChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args)
    {
        if (args.Reason == AutoSuggestionBoxTextChangeReason.UserInput) _ = UpdateSuggestionsAsync(delay: true);
    }

    private async void GlobalSearchBox_QuerySubmitted(AutoSuggestBox sender, AutoSuggestBoxQuerySubmittedEventArgs args)
    {
        if (args.ChosenSuggestion is ItemRef chosen)
        {
            await RevealAsync(chosen);
            return;
        }
        await SubmitSearchAsync();
    }

    private async void GlobalSearchButton_Click(object sender, RoutedEventArgs e) => await SubmitSearchAsync();

    /// <summary>Enter or the search button: go to the only (or exactly named) match, otherwise list the matches.</summary>
    private async Task SubmitSearchAsync()
    {
        var text = (GlobalSearchBox.Text ?? "").Trim();
        if (text.Length < 2)
        {
            ShowStatus("Type at least two characters to search.", InfoBarSeverity.Informational);
            return;
        }
        var results = await UpdateSuggestionsAsync(delay: false);
        if (results == null) return;
        var exact = results.Where(r => string.Equals(r.Name, text, StringComparison.CurrentCultureIgnoreCase)).ToList();
        if (results.Count == 1) await RevealAsync(results[0]);
        else if (exact.Count == 1) await RevealAsync(exact[0]);
        else if (results.Count == 0 && Guid.TryParse(text, out var id))
            ShowStatus($"Nothing in the term store has the ID {id}.", InfoBarSeverity.Informational);
        else if (results.Count == 0) ShowStatus($"Nothing in the term store starts with “{text}”.", InfoBarSeverity.Informational);
        else GlobalSearchBox.Focus(FocusState.Programmatic); // the suggestion list shows the matches
    }

    /// <summary>
    /// Fills the search box's list: term groups and loaded term sets whose name contains the text, term sets
    /// named exactly the text, and terms with a label or synonym starting with it (CSOM can't search term set
    /// names or labels by "contains"). Returns null when superseded or failed.
    /// </summary>
    private async Task<IReadOnlyList<ItemRef>?> UpdateSuggestionsAsync(bool delay)
    {
        _searchCts?.Cancel();
        var cts = _searchCts = new CancellationTokenSource();
        var text = (GlobalSearchBox.Text ?? "").Trim();
        if (text.Length < 2 || CurrentProfile is not { } profile || !_storeLoaded)
        {
            GlobalSearchBox.ItemsSource = null;
            return null;
        }

        try
        {
            if (delay) await Task.Delay(350, cts.Token); // wait for typing to pause

            // A unique identifier (with or without braces or dashes): look the item up by ID instead.
            if (Guid.TryParse(text, out var id))
            {
                var byId = await App.TermStore.FindByIdAsync(profile, id, cts.Token);
                if (cts.IsCancellationRequested) return null;
                GlobalSearchBox.ItemsSource = byId;
                GlobalSearchBox.IsSuggestionListOpen = byId.Count > 0 && GlobalSearchBox.FocusState != FocusState.Unfocused;
                return byId;
            }

            var local = AllNodes()
                .Where(n => n.Kind != TermNodeKind.Term && n.Name.Contains(text, StringComparison.CurrentCultureIgnoreCase))
                .Select(n => n.ToRef() with { GroupId = GroupOf(n)?.Id });
            var termSetsTask = App.TermStore.FindTermSetsByNameAsync(profile, text, cts.Token);
            var termsTask = App.TermStore.SearchTermsAsync(profile, text, null, 50, cts.Token);
            var termSets = await termSetsTask;
            var terms = await termsTask;
            if (cts.IsCancellationRequested) return null;

            var results = local.Concat(termSets)
                .GroupBy(r => (r.Kind, r.Id)).Select(g => g.First())
                .OrderBy(r => r.Kind)
                .ThenBy(r => r.Name, StringComparer.CurrentCultureIgnoreCase)
                .Concat(terms)
                .ToList();
            GlobalSearchBox.ItemsSource = results;
            GlobalSearchBox.IsSuggestionListOpen = results.Count > 0 && GlobalSearchBox.FocusState != FocusState.Unfocused;
            return results;
        }
        catch (OperationCanceledException)
        {
            return null;
        }
        catch (Exception ex) when (!cts.IsCancellationRequested)
        {
            GlobalSearchBox.ItemsSource = null;
            ShowError("Search failed", ex);
            return null;
        }
    }

    /// <summary>Opens the tree down to a search result, selects it and shows it in the editor.</summary>
    private async Task RevealAsync(ItemRef item)
    {
        if (CurrentProfile is not { } profile || _busy) return;
        if (!await ConfirmLeaveAsync()) return;
        GlobalSearchBox.IsSuggestionListOpen = false;
        ShowTermStore();

        SetBusy(true);
        TermNode? found = null;
        try
        {
            found = await FindInTreeAsync(profile, item);
        }
        catch (Exception ex)
        {
            ShowError($"Could not open the way to “{item.Name}”", ex);
        }
        finally
        {
            SetBusy(false);
        }
        if (found == null)
        {
            if (!StatusBar.IsOpen || StatusBar.Severity != InfoBarSeverity.Error)
                ShowStatus($"Couldn't find “{item.Name}” in the tree ({item.Path}). Try Refresh; site collection groups only show for their own site.", InfoBarSeverity.Warning);
            return;
        }

        TrySelectInTree(found);
        await ShowNodeAsync(found);
        await ScrollToMiddleAsync(found);
    }

    /// <summary>
    /// Scrolls the tree so the node's row is in the middle of the card (as far as the list allows near its top
    /// and bottom), rather than just inside the bottom edge. Waits briefly for the tree to create the row, which
    /// it does after the parents expand. The tree only creates rows near the visible part of the list, so a row
    /// further down is scrolled to through the tree's inner list first.
    /// </summary>
    private async Task ScrollToMiddleAsync(TermNode node)
    {
        for (int attempt = 0; attempt < 20; attempt++)
        {
            await Task.Delay(50);
            if (Tree.ContainerFromItem(node) is TreeViewItem row)
            {
                row.StartBringIntoView(new BringIntoViewOptions { VerticalAlignmentRatio = 0.5, AnimationDesired = true });
                return;
            }
            try
            {
                FindDescendant<ListViewBase>(Tree)?.ScrollIntoView(node);
            }
            catch (Exception)
            {
                // Not in the tree's list yet (its parents are still expanding); try again.
            }
        }
    }

    private static T? FindDescendant<T>(DependencyObject parent) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T match) return match;
            if (FindDescendant<T>(child) is { } found) return found;
        }
        return null;
    }

    /// <summary>Loads and expands the item's group, term set and parent terms, and returns its node.</summary>
    private async Task<TermNode?> FindInTreeAsync(ConnectionProfile profile, ItemRef item)
    {
        var existing = AllNodes().FirstOrDefault(n => n.Matches(item));
        if (existing != null)
        {
            Reveal(existing);
            return existing;
        }

        var group = Roots.FirstOrDefault(g => g.Id == (item.Kind == TermNodeKind.Group ? item.Id : item.GroupId));
        if (group == null || item.Kind == TermNodeKind.Group) return group;

        await OpenAsync(profile, group);
        var termSet = FindChild(group, t => t.Id == item.TermSetId);
        if (termSet == null || item.Kind == TermNodeKind.TermSet) return termSet;

        // Walk down the parent terms by default label.
        var current = termSet;
        foreach (var label in item.ParentLabels)
        {
            await OpenAsync(profile, current);
            var next = FindChild(current, c => string.Equals(c.Name, label, StringComparison.CurrentCultureIgnoreCase));
            if (next == null) break;
            current = next;
        }
        await OpenAsync(profile, current);
        var found = FindChild(current, c => c.Id == item.Id) ?? AllNodes().FirstOrDefault(n => n.Matches(item));
        if (found != null) Reveal(found);
        return found;
    }

    /// <summary>
    /// A loaded child of <paramref name="parent"/>, shown in the tree: if it was hidden, the tree shows the page
    /// around it instead of everything before it.
    /// </summary>
    private static TermNode? FindChild(TermNode parent, Func<TermNode, bool> match)
    {
        var child = parent.LoadedChildren.FirstOrDefault(match);
        if (child != null) parent.ShowPageOf(child);
        return child;
    }

    /// <summary>Expands the node's ancestors, showing the page around it (and around each ancestor) where hidden.</summary>
    private static void Reveal(TermNode node)
    {
        for (var child = node; child.Parent is { } parent; child = parent)
        {
            parent.ShowPageOf(child);
            parent.IsExpanded = true;
        }
    }

    private static async Task OpenAsync(ConnectionProfile profile, TermNode node)
    {
        if (node.HasUnrealizedChildren) await LoadChildrenAsync(profile, node);
        node.IsExpanded = true;
    }
}
