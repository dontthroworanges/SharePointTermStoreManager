using System.Diagnostics;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using SharePointTermStoreManager.Models;
using SharePointTermStoreManager.Services;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace SharePointTermStoreManager.Views;

/// <summary>
/// The Export page: choose a format, the term groups and whether to include deprecated terms, then save. The
/// term store is read fresh (<see cref="TermStoreService.ReadForExportAsync"/>) and written by
/// <see cref="ExportWriter"/>. The format and options are remembered.
/// </summary>
public sealed partial class ExportView : UserControl
{
    /// <summary>The term group check boxes, in tree order.</summary>
    private readonly List<(Guid Id, CheckBox Box)> _groups = [];

    private CancellationTokenSource? _cts;
    private string? _lastFile;

    public ExportView()
    {
        InitializeComponent();
        LoadOptions();
        foreach (var radio in FormatRadios) radio.Checked += (_, _) => SaveOptions();
        ResultBar.IsOpen = false;
    }

    private bool _loadingOptions;

    /// <summary>Shows the saved format and options (on start, and after Settings › Reset).</summary>
    public void LoadOptions()
    {
        _loadingOptions = true;
        var data = App.Settings.Data;
        var format = Enum.TryParse<ExportFormat>(data.ExportFormat, out var f) ? f : ExportFormat.Excel;
        foreach (var radio in FormatRadios) radio.IsChecked = (string)radio.Tag == format.ToString();
        IncludeDeprecatedBox.IsChecked = data.ExportIncludeDeprecated;
        _loadingOptions = false;
    }

    private RadioButton[] FormatRadios => [ExcelRadio, CsvRadio, JsonRadio, HtmlRadio];

    private ExportFormat SelectedFormat =>
        Enum.Parse<ExportFormat>((string)(FormatRadios.FirstOrDefault(r => r.IsChecked == true) ?? ExcelRadio).Tag);

    private bool Exporting => _cts != null;

    /// <summary>
    /// Follows the main window: whether a term store is loaded, and its term groups (new ones start checked;
    /// the others keep their check).
    /// </summary>
    public void UpdateState(bool loaded, IEnumerable<TermNode> groups)
    {
        _loaded = loaded;
        NotConnectedBar.IsOpen = !loaded;

        var roots = loaded ? groups.ToList() : [];
        if (!roots.Select(r => r.Id).SequenceEqual(_groups.Select(g => g.Id)))
        {
            var wasChecked = _groups.ToDictionary(g => g.Id, g => g.Box.IsChecked == true);
            _groups.Clear();
            GroupList.Children.Clear();
            foreach (var group in roots)
            {
                var label = group.IsSystemGroup ? $"{group.Name} (system)"
                    : group.IsSiteCollectionGroup ? $"{group.Name} (site collection)"
                    : group.Name;
                var box = new CheckBox { Content = label, IsChecked = !wasChecked.TryGetValue(group.Id, out var c) || c };
                box.Click += (_, _) => UpdateButtons();
                _groups.Add((group.Id, box));
                GroupList.Children.Add(box);
            }
        }
        NoGroupsText.Visibility = _groups.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        UpdateButtons();
    }

    private bool _loaded;

    private void UpdateButtons()
    {
        ExportButton.IsEnabled = _loaded && !Exporting && _groups.Any(g => g.Box.IsChecked == true);
        GroupCommands.Visibility = _groups.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        foreach (var (_, box) in _groups) box.IsEnabled = !Exporting;
        foreach (var radio in FormatRadios) radio.IsEnabled = !Exporting;
        IncludeDeprecatedBox.IsEnabled = !Exporting;
        CancelButton.Visibility = Exporting ? Visibility.Visible : Visibility.Collapsed;
        ProgressPanel.Visibility = CancelButton.Visibility;
    }

    private void SelectAllGroups_Click(object sender, RoutedEventArgs e) => CheckAll(true);

    private void ClearGroups_Click(object sender, RoutedEventArgs e) => CheckAll(false);

    private void CheckAll(bool check)
    {
        foreach (var (_, box) in _groups) box.IsChecked = check;
        UpdateButtons();
    }

    private void Option_Click(object sender, RoutedEventArgs e) => SaveOptions();

    private void SaveOptions()
    {
        if (_loadingOptions) return;
        App.Settings.Data.ExportFormat = SelectedFormat.ToString();
        App.Settings.Data.ExportIncludeDeprecated = IncludeDeprecatedBox.IsChecked == true;
        App.Settings.Save();
    }

    // ---- Exporting -----------------------------------------------------------------------------------------

    private async void Export_Click(object sender, RoutedEventArgs e)
    {
        if (App.Settings.CurrentProfile is not { } profile || Exporting) return;
        var format = SelectedFormat;
        bool includeDeprecated = IncludeDeprecatedBox.IsChecked == true;
        var chosen = _groups.Where(g => g.Box.IsChecked == true).Select(g => g.Id).ToHashSet();
        // Everything checked exports every group, including any added since the tree was loaded.
        IReadOnlySet<Guid>? groupIds = chosen.Count == _groups.Count ? null : chosen;

        var picker = new FileSavePicker
        {
            SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
            SuggestedFileName = SafeFileName($"{profile.EffectiveName} term store {DateTime.Now:yyyy-MM-dd}"),
        };
        picker.FileTypeChoices.Add(ExportWriter.FileTypeName(format), [ExportWriter.Extension(format)]);
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(App.MainWindow));
        if (await picker.PickSaveFileAsync() is not { } file) return;
        var path = file.Path;
        var temp = path + ".tmp";

        ResultBar.IsOpen = false;
        _cts = new CancellationTokenSource();
        UpdateButtons();
        ExportProgressBar.IsIndeterminate = true;
        ProgressText.Text = "Connecting…";
        var progress = new Progress<ExportProgress>(p =>
        {
            if (p.Total > 0)
            {
                ExportProgressBar.IsIndeterminate = false;
                ExportProgressBar.Maximum = p.Total;
                ExportProgressBar.Value = p.Done;
                ProgressText.Text = p.Done < p.Total ? $"{p.Text} (term set {p.Done + 1:N0} of {p.Total:N0})" : p.Text;
            }
            else
            {
                ProgressText.Text = p.Text;
            }
        });

        try
        {
            var token = _cts.Token;
            // On the UI thread: signing in again, if the session has expired, shows a window.
            var store = await App.TermStore.ReadForExportAsync(profile, groupIds, includeDeprecated, progress, token);
            ExportProgressBar.IsIndeterminate = true;
            ProgressText.Text = "Writing the file…";
            await Task.Run(() =>
            {
                ExportWriter.Write(store, format, temp);
                File.Move(temp, path, overwrite: true);
            }, token);

            _lastFile = path;
            ShowResult(InfoBarSeverity.Success, "Export complete",
                $"Exported {Count(store.Groups.Count, "term group")}, {Count(store.TermSetCount, "term set")} and " +
                $"{Count(store.TermCount, "term")} to {path}.", withActions: true);
        }
        catch (OperationCanceledException)
        {
            Cleanup(temp, path);
            ShowResult(InfoBarSeverity.Informational, "Export cancelled", "Nothing was saved.", withActions: false);
        }
        catch (Exception ex)
        {
            Cleanup(temp, path);
            ShowResult(InfoBarSeverity.Error, "Export failed", ex.Message, withActions: false);
        }
        finally
        {
            _cts.Dispose();
            _cts = null;
            UpdateButtons();
        }
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        _cts?.Cancel();
        ProgressText.Text = "Cancelling…";
    }

    /// <summary>Removes the partly written file, and the empty file the save dialog made.</summary>
    private static void Cleanup(string temp, string path)
    {
        try
        {
            if (File.Exists(temp)) File.Delete(temp);
            if (File.Exists(path) && new FileInfo(path).Length == 0) File.Delete(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Debug.WriteLine($"Could not clean up after the export: {ex.Message}");
        }
    }

    private void ShowResult(InfoBarSeverity severity, string title, string message, bool withActions)
    {
        ResultBar.Severity = severity;
        ResultBar.Title = title;
        ResultBar.Message = message;
        ResultActions.Visibility = withActions ? Visibility.Visible : Visibility.Collapsed;
        ResultBar.IsOpen = true;
    }

    private void OpenFile_Click(object sender, RoutedEventArgs e)
    {
        if (_lastFile == null) return;
        try
        {
            Process.Start(new ProcessStartInfo(_lastFile) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            ShowResult(InfoBarSeverity.Error, "Could not open the file", ex.Message, withActions: true);
        }
    }

    private void ShowInFolder_Click(object sender, RoutedEventArgs e)
    {
        if (_lastFile == null) return;
        try
        {
            Process.Start("explorer.exe", $"/select,\"{_lastFile}\"");
        }
        catch (Exception ex)
        {
            ShowResult(InfoBarSeverity.Error, "Could not open the folder", ex.Message, withActions: true);
        }
    }

    private static string Count(int count, string noun) => count == 1 ? $"1 {noun}" : $"{count:N0} {noun}s";

    private static string SafeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return new string(name.Select(c => invalid.Contains(c) ? '-' : c).ToArray()).Trim();
    }
}
