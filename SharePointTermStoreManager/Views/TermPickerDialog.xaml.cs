using System.Collections.ObjectModel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using SharePointTermStoreManager.Models;

namespace SharePointTermStoreManager.Views;

/// <summary>What a <see cref="TermPickerDialog"/> asks for.</summary>
public sealed class PickerOptions
{
    public string Title { get; init; } = "Choose";
    public string Intro { get; init; } = "";
    public string PrimaryButtonText { get; init; } = "Select";

    /// <summary>Kinds of item that can be picked.</summary>
    public IReadOnlySet<TermNodeKind> Allowed { get; init; } = new HashSet<TermNodeKind> { TermNodeKind.Term };

    /// <summary>The deepest level the browse tree opens to, e.g. Group to pick only groups.</summary>
    public TermNodeKind Deepest { get; init; } = TermNodeKind.Term;

    /// <summary>Offers searching just this term set, and starts with it selected.</summary>
    public ItemRef? ScopeTermSet { get; init; }

    /// <summary>Extra rules for the picked item; returns why it can't be used, or null.</summary>
    public Func<ItemRef, string?>? Validate { get; init; }
}

/// <summary>
/// Picks a term group, term set or term: search labels and synonyms (when terms can be picked) or browse the term
/// store. Changes nothing; the caller acts on <see cref="Result"/>.
/// </summary>
public sealed partial class TermPickerDialog : ContentDialog
{
    private readonly ConnectionProfile _profile;
    private readonly PickerOptions _options;
    private readonly ObservableCollection<TermNode> _roots = new();
    private CancellationTokenSource? _searchCts;
    private ItemRef? _candidate;

    public TermPickerDialog(ConnectionProfile profile, PickerOptions options)
    {
        InitializeComponent();
        _profile = profile;
        _options = options;

        Title = options.Title;
        PrimaryButtonText = options.PrimaryButtonText;
        IntroText.Text = options.Intro;
        IntroText.Visibility = options.Intro.Length > 0 ? Visibility.Visible : Visibility.Collapsed;
        BrowseTree.ItemsSource = _roots;

        bool canSearch = options.Allowed.Contains(TermNodeKind.Term);
        SearchRow.Visibility = canSearch ? Visibility.Visible : Visibility.Collapsed;
        ScopeBox.Items.Add("Whole term store");
        if (options.ScopeTermSet is { } scope) ScopeBox.Items.Add($"In “{scope.Name}”");
        // Start in the term set unless Settings says to search the whole term store first.
        ScopeBox.SelectedIndex = App.Settings.Data.PickerSearchScope == PickerScopes.WholeStore ? 0 : ScopeBox.Items.Count - 1;
        ScopeBox.Visibility = ScopeBox.Items.Count > 1 ? Visibility.Visible : Visibility.Collapsed;

        Opened += async (_, _) =>
        {
            if (canSearch) SearchBox.Focus(FocusState.Programmatic);
            await LoadRootsAsync();
        };
        Closed += (_, _) => _searchCts?.Cancel();
    }

    /// <summary>The picked item, once the dialog is accepted.</summary>
    public ItemRef? Result { get; private set; }

    // ---- Browse --------------------------------------------------------------------------------------------

    private async Task LoadRootsAsync()
    {
        LoadingRing.IsActive = true;
        try
        {
            var store = await App.TermStore.GetTermStoreAsync(_profile);
            foreach (var group in store.Groups)
            {
                if (_options.Deepest == TermNodeKind.Group) group.HasUnrealizedChildren = false;
                _roots.Add(group);
            }
        }
        catch (Exception ex)
        {
            EmptyText.Text = $"Could not load the term store: {ex.Message}";
        }
        finally
        {
            LoadingRing.IsActive = false;
        }
    }

    private async void BrowseTree_Expanding(TreeView sender, TreeViewExpandingEventArgs args)
    {
        if (args.Item is not TermNode node || !node.HasUnrealizedChildren || node.IsLoading) return;
        node.IsLoading = true;
        try
        {
            var children = await App.TermStore.GetChildrenAsync(_profile, node);
            node.Children.Clear();
            foreach (var child in children)
            {
                // Stop at the deepest level asked for, but terms nest under terms, so a term still opens to its child terms.
                if (child.Kind == _options.Deepest && child.Kind != TermNodeKind.Term) child.HasUnrealizedChildren = false;
                node.Children.Add(child);
            }
            node.HasUnrealizedChildren = false;
        }
        catch (Exception ex)
        {
            args.Node.IsExpanded = false;
            ShowProblem($"Could not load “{node.Name}”: {ex.Message}");
        }
        finally
        {
            node.IsLoading = false;
        }
    }

    private void BrowseTree_ItemInvoked(TreeView sender, TreeViewItemInvokedEventArgs args)
    {
        if (args.InvokedItem is TermNode node) Consider(node.ToRef());
    }

    // ---- Search --------------------------------------------------------------------------------------------

    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e) => _ = SearchAsync();

    private void ScopeBox_SelectionChanged(object sender, SelectionChangedEventArgs e) => _ = SearchAsync();

    private async Task SearchAsync()
    {
        _searchCts?.Cancel();
        var cts = _searchCts = new CancellationTokenSource();
        var text = (SearchBox.Text ?? "").Trim();

        bool searching = text.Length >= 2;
        BrowseTree.Visibility = searching ? Visibility.Collapsed : Visibility.Visible;
        ResultsList.Visibility = searching ? Visibility.Visible : Visibility.Collapsed;
        EmptyText.Text = text.Length == 1 ? "Type at least two characters to search." : "";
        if (!searching)
        {
            ResultsList.ItemsSource = null;
            LoadingRing.IsActive = false;
            return;
        }

        try
        {
            await Task.Delay(350, cts.Token); // wait for typing to pause
            LoadingRing.IsActive = true;
            Guid? scope = ScopeBox.SelectedIndex == 1 ? _options.ScopeTermSet?.Id : null;
            var results = await App.TermStore.SearchTermsAsync(_profile, text, scope, ct: cts.Token);
            if (cts.IsCancellationRequested) return;
            ResultsList.ItemsSource = results;
            EmptyText.Text = results.Count == 0 ? $"No term names or synonyms start with “{text}”." : "";
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex) when (!cts.IsCancellationRequested)
        {
            ResultsList.ItemsSource = null;
            EmptyText.Text = $"Search failed: {ex.Message}";
        }
        finally
        {
            if (!cts.IsCancellationRequested) LoadingRing.IsActive = false;
        }
    }

    private void ResultsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ResultsList.SelectedItem is ItemRef item) Consider(item);
    }

    // ---- Choosing ------------------------------------------------------------------------------------------

    private void Consider(ItemRef item)
    {
        _candidate = item;
        SelectionText.Text = $"Selected: {item.Path}";

        string? problem = !_options.Allowed.Contains(item.Kind)
            ? $"Pick {string.Join(" or ", _options.Allowed.Select(k => "a " + TermNode.Noun(k)))}."
            : _options.Validate?.Invoke(item);
        if (problem != null) ShowProblem(problem);
        else ProblemText.Visibility = Visibility.Collapsed;
        IsPrimaryButtonEnabled = problem == null;
    }

    private void ShowProblem(string problem)
    {
        ProblemText.Text = problem;
        ProblemText.Visibility = Visibility.Visible;
    }

    private void ContentDialog_PrimaryButtonClick(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        if (_candidate == null)
        {
            args.Cancel = true;
            return;
        }
        Result = _candidate;
    }
}
