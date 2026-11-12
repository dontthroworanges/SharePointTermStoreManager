using System.Diagnostics;
using System.Globalization;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using SharePointTermStoreManager.Helpers;
using SharePointTermStoreManager.Models;
using SharePointTermStoreManager.Services;
using Windows.ApplicationModel.DataTransfer;

namespace SharePointTermStoreManager.Views;

/// <summary>
/// The Operations Log page: the entries in the monthly log files, newest first, narrowed by a search box and
/// filters (date, action, outcome, term group and, when there's more than one, site). Each entry shows a summary
/// and expands to its JSON, pretty-printed (see <see cref="JsonView"/>). The page reads the files again when it's
/// shown after something was logged, and straight away when something is logged while it's open.
/// </summary>
public sealed partial class OperationsLogView : UserControl
{
    private IReadOnlyList<LogEntry> _all = [];
    private List<LogEntry> _shown = [];
    private int _skipped;

    /// <summary>True when the files have changed since they were last read (or were never read).</summary>
    private bool _stale = true;

    private bool _loading;

    /// <summary>True while filter controls are being changed from code, so each change doesn't filter again.</summary>
    private bool _updating;

    private readonly bool _ready;

    /// <summary>Whether new entries start expanded (set by Expand all / Collapse all).</summary>
    private bool _expandByDefault = true;

    /// <summary>Filters the list when typing in the search box pauses.</summary>
    private readonly DispatcherQueueTimer _searchTimer;

    public OperationsLogView()
    {
        InitializeComponent();

        _searchTimer = DispatcherQueue.CreateTimer();
        _searchTimer.Interval = TimeSpan.FromMilliseconds(250);
        _searchTimer.IsRepeating = false;
        _searchTimer.Tick += (_, _) => ApplyFilters();

        OperationLog.Changed += (_, _) => DispatcherQueue.TryEnqueue(() =>
        {
            _stale = true;
            if (Visibility == Visibility.Visible) _ = ReloadAsync();
        });

        _updating = true;
        SetChoices(ActionBox, "Any action", []);
        SetChoices(GroupBox, "Any term group", []);
        SetChoices(SiteBox, "Any site", []);
        _updating = false;
        _ready = true;
    }

    /// <summary>Called when the page is shown: reads the log files if they changed since the last read.</summary>
    public void Activate()
    {
        if (_stale) _ = ReloadAsync();
    }

    // ---- Loading -------------------------------------------------------------------------------------------

    private async Task ReloadAsync()
    {
        if (_loading)
        {
            _stale = true; // read again when the current read finishes
            return;
        }
        _loading = true;
        LoadingRing.IsActive = true;
        try
        {
            do
            {
                _stale = false;
                var (entries, skipped) = await Task.Run(() => (OperationLog.ReadAll(out int s), s));

                // Keep entries open or closed as they were; new ones follow Expand all / Collapse all.
                var collapsed = _all.Where(e => e.IsExpanded != _expandByDefault).Select(e => e.Raw).ToHashSet();
                foreach (var entry in entries)
                    entry.IsExpanded = collapsed.Contains(entry.Raw) ? !_expandByDefault : _expandByDefault;

                _all = entries;
                _skipped = skipped;
            }
            while (_stale);
        }
        catch (Exception ex)
        {
            App.MainWindow.ShowStatus($"Could not read the operations log: {ex.Message}", InfoBarSeverity.Error);
        }
        finally
        {
            _loading = false;
            LoadingRing.IsActive = false;
        }

        _updating = true;
        SetChoices(ActionBox, "Any action", _all.Select(e => e.OperationText));
        SetChoices(GroupBox, "Any term group", _all.SelectMany(e => e.Groups));
        var sites = _all.Select(e => e.Site).Where(s => s.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        SetChoices(SiteBox, "Any site", sites);
        SitePanel.Visibility = sites.Count > 1 || SiteBox.SelectedIndex > 0 ? Visibility.Visible : Visibility.Collapsed;
        _updating = false;
        ApplyFilters(keepScroll: true);
    }

    /// <summary>
    /// Fills a filter's list with "any" and the distinct values, sorted, keeping the selection when it's still
    /// there.
    /// </summary>
    private static void SetChoices(ComboBox box, string any, IEnumerable<string> values)
    {
        var previous = box.SelectedIndex > 0 ? box.SelectedItem as string : null;
        box.Items.Clear();
        box.Items.Add(any);
        foreach (var value in values.Where(v => v.Length > 0).Distinct(StringComparer.CurrentCultureIgnoreCase)
                     .Order(StringComparer.CurrentCultureIgnoreCase))
            box.Items.Add(value);

        int index = 0;
        for (int i = 1; i < box.Items.Count; i++)
        {
            if (string.Equals(box.Items[i] as string, previous, StringComparison.CurrentCultureIgnoreCase)) index = i;
        }
        box.SelectedIndex = index;
    }

    // ---- Filtering -----------------------------------------------------------------------------------------

    private void ApplyFilters(bool keepScroll = false)
    {
        if (!_ready || _updating) return;
        _searchTimer.Stop();

        var (from, to) = DateRange();
        var action = Selected(ActionBox);
        var group = Selected(GroupBox);
        var site = Selected(SiteBox);
        string? outcome = OutcomeBox.SelectedIndex switch
        {
            1 => "succeeded",
            2 => "failed",
            3 => "started",
            _ => null,
        };
        var phrase = (SearchBox.Text ?? "").Trim();

        _shown = _all.Where(e =>
                (from == null || e.Time.ToLocalTime().Date >= from) &&
                (to == null || e.Time.ToLocalTime().Date <= to) &&
                (action == null || e.OperationText == action) &&
                (outcome == null || string.Equals(e.Outcome, outcome, StringComparison.OrdinalIgnoreCase)) &&
                (group == null || e.Groups.Contains(group, StringComparer.CurrentCultureIgnoreCase)) &&
                (site == null || string.Equals(e.Site, site, StringComparison.OrdinalIgnoreCase)) &&
                (phrase.Length == 0 || e.SearchText.Contains(phrase, StringComparison.CurrentCultureIgnoreCase)))
            .ToList();

        Entries.ItemsSource = _shown;
        if (!keepScroll) EntriesScroller.ChangeView(null, 0, null, disableAnimation: true);

        bool filtered = FiltersActive;
        ClearFiltersButton.IsEnabled = filtered;
        CountText.Text = _all.Count == 0 ? ""
            : filtered ? $"{_shown.Count:N0} of {CountLabel(_all.Count)}"
            : CountLabel(_all.Count);
        if (_skipped > 0) CountText.Text += $" · {_skipped:N0} unreadable {(_skipped == 1 ? "line" : "lines")} skipped";

        CopyButton.IsEnabled = ExpandAllButton.IsEnabled = CollapseAllButton.IsEnabled = _shown.Count > 0;

        if (_all.Count == 0)
        {
            EmptyIcon.Glyph = FluentIcons.History;
            EmptyText.Text = "Nothing has been logged yet. Every change the app makes to the term store is recorded here: new items, edits (names, descriptions, synonyms, properties and settings), moves, merges, copies, reuse, deprecations and deletes.";
        }
        else
        {
            EmptyIcon.Glyph = FluentIcons.FilterDismiss;
            EmptyText.Text = "No entries match the search and filters.";
        }
        EmptyState.Visibility = _shown.Count == 0 && !_loading ? Visibility.Visible : Visibility.Collapsed;
    }

    private static string CountLabel(int count) => count == 1 ? "1 entry" : $"{count:N0} entries";

    private static string? Selected(ComboBox box) => box.SelectedIndex > 0 ? box.SelectedItem as string : null;

    private bool FiltersActive =>
        !string.IsNullOrWhiteSpace(SearchBox.Text) || DateBox.SelectedIndex > 0 || ActionBox.SelectedIndex > 0 ||
        OutcomeBox.SelectedIndex > 0 || GroupBox.SelectedIndex > 0 || SiteBox.SelectedIndex > 0;

    /// <summary>The first and last day to show (local dates, inclusive); null for no limit.</summary>
    private (DateTime? From, DateTime? To) DateRange()
    {
        var today = DateTime.Today;
        return DateBox.SelectedIndex switch
        {
            1 => (today, today),
            2 => (today.AddDays(-6), null),
            3 => (today.AddDays(-29), null),
            4 => (new DateTime(today.Year, today.Month, 1), null),
            5 => (FromDate.Date?.Date, ToDate.Date?.Date),
            _ => (null, null),
        };
    }

    private void Filter_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (!_ready) return;
        if (ReferenceEquals(sender, DateBox))
        {
            var custom = DateBox.SelectedIndex == 5 ? Visibility.Visible : Visibility.Collapsed;
            FromPanel.Visibility = ToPanel.Visibility = custom;
        }
        ApplyFilters();
    }

    private void Date_Changed(CalendarDatePicker sender, CalendarDatePickerDateChangedEventArgs args) => ApplyFilters();

    private void SearchBox_TextChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args)
    {
        _searchTimer.Stop();
        _searchTimer.Start();
    }

    private void ClearFilters_Click(object sender, RoutedEventArgs e)
    {
        _updating = true;
        SearchBox.Text = "";
        DateBox.SelectedIndex = 0;
        FromDate.Date = null;
        ToDate.Date = null;
        ActionBox.SelectedIndex = OutcomeBox.SelectedIndex = GroupBox.SelectedIndex = SiteBox.SelectedIndex = 0;
        FromPanel.Visibility = ToPanel.Visibility = Visibility.Collapsed;
        _updating = false;
        ApplyFilters();
    }

    // ---- Commands ------------------------------------------------------------------------------------------

    private void Refresh_Click(object sender, RoutedEventArgs e) => _ = ReloadAsync();

    private void ExpandAll_Click(object sender, RoutedEventArgs e) => SetAllExpanded(true);

    private void CollapseAll_Click(object sender, RoutedEventArgs e) => SetAllExpanded(false);

    private void SetAllExpanded(bool expanded)
    {
        _expandByDefault = expanded;
        foreach (var entry in _all) entry.IsExpanded = expanded;
    }

    private void CopyJsonLines_Click(object sender, RoutedEventArgs e) =>
        Copy(string.Join(Environment.NewLine, _shown.Select(x => x.Raw)) + Environment.NewLine, _shown.Count, "as JSON Lines");

    private void CopyFormatted_Click(object sender, RoutedEventArgs e) =>
        Copy(LogEntry.FormatAll(_shown), _shown.Count, "as formatted JSON");

    private void CopyEntry_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.Tag is LogEntry entry) Copy(entry.FormattedJson, 1, "as formatted JSON");
    }

    private static void Copy(string text, int count, string format)
    {
        if (count == 0) return;
        var package = new DataPackage();
        package.SetText(text);
        Clipboard.SetContent(package);
        App.MainWindow.ShowStatus($"Copied {CountLabel(count)} to the clipboard {format}.", InfoBarSeverity.Success);
    }

    /// <summary>Lists the log files (one per month, newest first) and "Show in folder".</summary>
    private void OpenFileMenu_Opening(object sender, object e)
    {
        OpenFileMenu.Items.Clear();
        var files = OperationLog.Files();
        if (files.Count == 0)
            OpenFileMenu.Items.Add(new MenuFlyoutItem { Text = "No log files yet", IsEnabled = false });

        foreach (var file in files)
        {
            var name = Path.GetFileName(file);
            var item = new MenuFlyoutItem
            {
                Text = $"{MonthOf(name)}  ({name})",
                Icon = FluentIcons.Create(FluentIcons.DocumentText),
            };
            ToolTipService.SetToolTip(item, file);
            item.Click += (_, _) => Launch(file);
            OpenFileMenu.Items.Add(item);
        }

        OpenFileMenu.Items.Add(new MenuFlyoutSeparator());
        var folder = new MenuFlyoutItem { Text = "Show in folder", Icon = FluentIcons.Create(FluentIcons.FolderOpen) };
        folder.Click += (_, _) => ShowInFolder(files.FirstOrDefault());
        OpenFileMenu.Items.Add(folder);
    }

    /// <summary>"November 2026" for "operations-2026-11.jsonl".</summary>
    private static string MonthOf(string fileName)
    {
        var stem = Path.GetFileNameWithoutExtension(fileName).Replace("operations-", "");
        return DateTime.TryParseExact(stem, "yyyy-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out var month)
            ? month.ToString("MMMM yyyy", CultureInfo.CurrentCulture)
            : stem;
    }

    /// <summary>Opens the file in the app Windows uses for .jsonl files (Windows asks which, the first time).</summary>
    private static void Launch(string file)
    {
        try
        {
            Process.Start(new ProcessStartInfo(file) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            App.MainWindow.ShowStatus($"Could not open {Path.GetFileName(file)}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private static void ShowInFolder(string? file)
    {
        try
        {
            if (file != null && File.Exists(file))
            {
                Process.Start("explorer.exe", $"/select,\"{file}\"");
            }
            else
            {
                Directory.CreateDirectory(OperationLog.Folder);
                Process.Start(new ProcessStartInfo(OperationLog.Folder) { UseShellExecute = true });
            }
        }
        catch (Exception ex)
        {
            App.MainWindow.ShowStatus($"Could not open the log folder: {ex.Message}", InfoBarSeverity.Error);
        }
    }
}
