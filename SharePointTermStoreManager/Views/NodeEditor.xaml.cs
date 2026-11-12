using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using SharePointTermStoreManager.Helpers;
using SharePointTermStoreManager.Models;
using Windows.System;

namespace SharePointTermStoreManager.Views;

/// <summary>One editable row in a custom property list.</summary>
public sealed partial class PropertyRow : ObservableObject
{
    [ObservableProperty]
    private string _key = "";

    [ObservableProperty]
    private string _value = "";

    [ObservableProperty]
    private bool _isEditable = true;
}

/// <summary>
/// Shows and edits the selected term group, term set or term. The editor only collects changes: the window saves
/// them (<see cref="SaveRequested"/>) and then calls <see cref="Show"/> with what SharePoint now holds.
/// </summary>
public sealed partial class NodeEditor : UserControl
{
    private readonly ObservableCollection<PropertyRow> _customProps = new();
    private readonly ObservableCollection<PropertyRow> _localProps = new();

    /// <summary>The term's synonyms as shown in the chips (SharePoint's spelling for ones it already has).</summary>
    private readonly List<string> _synonyms = new();
    private readonly List<Button> _synonymRemoveButtons = new();

    private NodeDetails? _original;
    private bool _readOnly;
    private bool _busy;
    private bool _loadingForm;

    public NodeEditor()
    {
        InitializeComponent();
        CustomPropsList.ItemsSource = _customProps;
        LocalPropsList.ItemsSource = _localProps;
        _customProps.CollectionChanged += Rows_CollectionChanged;
        _localProps.CollectionChanged += Rows_CollectionChanged;
    }

    /// <summary>The node shown, or null.</summary>
    public TermNode? Node { get; private set; }

    /// <summary>The details as last read from SharePoint, or null while loading.</summary>
    public NodeDetails? Original => _original;

    /// <summary>True when the form differs from what was read from SharePoint.</summary>
    public bool IsDirty { get; private set; }

    public event EventHandler? SaveRequested;
    public event EventHandler? DirtyChanged;

    /// <summary>An ancestor was clicked in the breadcrumb; the window selects it.</summary>
    public event EventHandler<TermNode>? PathItemInvoked;

    // ---- Showing -------------------------------------------------------------------------------------------

    /// <summary>
    /// Shows the "nothing selected" state: the placeholder heading, and the command bar with the placeholder
    /// commands (see <see cref="SetPlaceholderCommands"/>), all disabled.
    /// </summary>
    public void Clear()
    {
        Node = null;
        _original = null;
        _readOnly = false;
        SetTitle("", null);
        LoadingRing.IsActive = false;
        HintBar.IsOpen = false;
        ValidationBar.IsOpen = false;
        UpdateHintPanel();
        FormScroller.Visibility = Visibility.Collapsed;
        SaveButton.Visibility = DiscardButton.Visibility = Visibility.Visible;
        if (_placeholderCommands != null) SetItemCommands(_placeholderCommands());
        ItemCommands.Visibility = Vis(_placeholderCommands != null);
        SetDirty(false);
    }

    private Func<IEnumerable<ICommandBarElement>>? _placeholderCommands;

    /// <summary>
    /// Sets the commands the bar shows (disabled) while no item is selected, as a preview of what an item
    /// offers. A factory, because a command bar element can only be in one bar at a time.
    /// </summary>
    public void SetPlaceholderCommands(Func<IEnumerable<ICommandBarElement>> commands)
    {
        _placeholderCommands = commands;
        if (Node == null) Clear();
    }

    /// <summary>Shows the node's name while its details are read.</summary>
    public void ShowLoading(TermNode node)
    {
        Clear();
        Node = node;
        SetTitle(node.Name, node);
        LoadingRing.IsActive = true;
    }

    /// <summary>Fills the form with <paramref name="details"/> and discards any edits.</summary>
    public void Show(TermNode node, NodeDetails details)
    {
        Node = node;
        _original = details;
        ValidationBar.IsOpen = false;
        LoadingRing.IsActive = false;

        SetTitle(details.Name, node);

        bool isTerm = node.Kind == TermNodeKind.Term;
        bool isTermSet = node.Kind == TermNodeKind.TermSet;
        bool hasUsage = isTerm || isTermSet;
        string language = LanguageName(details.Language);

        // General: translations and synonyms (terms) or name and description (groups and term sets).
        GeneralHeading.Text = isTerm ? "Translations and synonyms" : "Name and description";
        LanguageText.Text = isTerm
            ? $"Labels and description in {language}, the term store's default language."
            : $"In {language}, the term store's default language.";
        NameBox.Header = isTerm ? "Default label" : "Name";
        SynonymsPanel.Visibility = Vis(isTerm);
        OtherLanguagesPanel.Visibility = Vis(isTerm && details.OtherLanguages.Count > 0);
        OtherLanguagesList.ItemsSource = details.OtherLanguages;
        TermSetPanel.Visibility = Vis(isTermSet);

        // Usage settings.
        TaggingPanel.Visibility = Vis(hasUsage);
        TaggingCaption.Text = isTerm
            ? "When enabled, this term is available to end users and content editors of sites using this term set."
            : "When enabled, the terms in this term set are available to end users and content editors for tagging.";
        SubmissionPanel.Visibility = Vis(isTermSet);
        SortOrderPanel.Visibility = Vis(hasUsage);
        SortOrderCaption.Text = isTerm
            ? "Child terms can be sorted alphabetically or in a custom order."
            : "Terms can be sorted alphabetically or in a custom order.";
        SortOrderText.Text = details.HasCustomSortOrder ? "Custom order" : "Alphabetical";
        MemberOfPanel.Visibility = Vis(isTerm && details.MemberOf.Count > 0);
        MemberOfList.ItemsSource = details.MemberOf;
        StatusCaption.Visibility = StatusText.Visibility = Vis(isTerm);

        // Advanced: custom properties (term sets and terms only).
        AdvancedSection.Visibility = Vis(hasUsage);
        CustomPropsHeading.Text = isTerm ? "Shared custom properties" : "Custom properties";
        CustomPropsCaption.Text = isTerm
            ? "Shared custom properties are available on all reused or pinned instances of this term anywhere in the term store."
            : "Key/value pairs stored on the term set, e.g. for navigation or integrations.";
        LocalPropsPanel.Visibility = Vis(isTerm);

        _loadingForm = true;
        NameBox.Text = details.Name;
        DescriptionBox.Text = details.Description;
        _synonyms.Clear();
        _synonyms.AddRange(details.Synonyms);
        SynonymInput.Text = "";
        SynonymError.Visibility = Visibility.Collapsed;
        RebuildSynonymChips();
        OwnerBox.Text = details.Owner;
        ContactBox.Text = details.Contact;
        OpenBox.IsOn = details.IsOpenForTermCreation;
        TaggingBox.IsOn = details.IsAvailableForTagging;
        FillRows(_customProps, details.CustomProperties);
        FillRows(_localProps, details.LocalCustomProperties);
        _loadingForm = false;

        PathText.Text = node.Path;
        IdText.Text = details.Id.ToString();
        NameIdText.Text = $"{details.Name}|{details.Id}";
        NameIdPanel.Visibility = Vis(isTerm);
        CreatedText.Text = FormatDate(details.Created);
        ModifiedText.Text = FormatDate(details.LastModified);
        StatusText.Text = StatusOf(details);

        // Read-only items, with the reason.
        string? hint = null;
        var severity = InfoBarSeverity.Informational;
        _readOnly = false;
        if (node.Kind == TermNodeKind.Group && details.IsSystemGroup)
        {
            _readOnly = true;
            hint = "System term groups are managed by SharePoint and can't be renamed or deleted. Their term sets can be edited.";
        }
        else if (isTerm && details.IsPinned && !details.IsSourceTerm)
        {
            _readOnly = true;
            hint = "This is a pinned copy of a term. Pinned copies can only be changed in the source term set.";
        }
        else if (isTerm && details.IsDeprecated)
        {
            hint = "This term is deprecated: it stays in the term store and on content already tagged with it, but can't be used for new tags. Un-deprecate in the command bar makes it available again.";
        }
        else if (isTerm && details.IsReused)
        {
            severity = InfoBarSeverity.Warning;
            hint = "This term is reused in other term sets. Its labels, description and custom properties are shared, so changes to them show everywhere it's used.";
        }
        HintBar.Message = hint ?? "";
        HintBar.Severity = severity;
        HintBar.IsOpen = hint != null;
        UpdateHintPanel();

        FormScroller.Visibility = Visibility.Visible;
        SaveButton.Visibility = DiscardButton.Visibility = Vis(!_readOnly);
        UpdateSeparator();
        ItemCommands.Visibility = Visibility.Visible;
        UpdateEnabled();
        SetDirty(false);
    }

    /// <summary>Disables editing while a save is running.</summary>
    public void SetBusy(bool busy)
    {
        _busy = busy;
        LoadingRing.IsActive = busy;
        UpdateEnabled();
    }

    /// <summary>Moves keyboard focus to the name box, e.g. for "Rename".</summary>
    public void FocusName()
    {
        if (FormScroller.Visibility != Visibility.Visible || _readOnly) return;
        NameBox.Focus(FocusState.Programmatic);
        NameBox.SelectAll();
    }

    // ---- Command bar ---------------------------------------------------------------------------------------

    /// <summary>
    /// Replaces the item's commands in the command bar, after Save and Discard. The bar shows them all when
    /// there's room and moves the last ones into its "…" menu (shown only then) when the panel is narrow.
    /// </summary>
    public void SetItemCommands(IEnumerable<ICommandBarElement> commands)
    {
        // Save, Discard and the separator after them come first and stay.
        while (ItemCommands.PrimaryCommands.Count > 3) ItemCommands.PrimaryCommands.RemoveAt(3);
        foreach (var command in commands) ItemCommands.PrimaryCommands.Add(command);
        UpdateSeparator();
    }

    /// <summary>Disables every command while an operation is running.</summary>
    public void SetCommandsEnabled(bool enabled) => ItemCommands.IsEnabled = enabled;

    /// <summary>The separator only shows between Save/Discard and the item's commands when both are there.</summary>
    private void UpdateSeparator() =>
        SaveSeparator.Visibility = Vis(!_readOnly && ItemCommands.PrimaryCommands.Count > 3);

    // ---- Reading the form ----------------------------------------------------------------------------------

    /// <summary>
    /// The edited details, normalized as SharePoint stores them, or null with <paramref name="error"/> set when
    /// something must be fixed first.
    /// </summary>
    public NodeDetails? GetEdited(out string? error)
    {
        error = null;
        if (_original is not { } original) return null;

        // Text that only differs from SharePoint's by normalization (spacing, line endings) keeps SharePoint's
        // value, so it doesn't count as a change and isn't rewritten on save.
        var name = KeepOriginal(TermNames.Normalize(NameBox.Text), original.Name, TermNames.Normalize);
        var description = KeepOriginal(NormalizeDescription(DescriptionBox.Text), original.Description, NormalizeDescription);

        // The chips, plus anything typed in the synonym box but not added yet, so it isn't lost on save.
        var synonyms = new List<string>();
        foreach (var entry in _synonyms.Concat(SplitSynonyms(SynonymInput.Text)))
        {
            var synonym = TermNames.Normalize(entry);
            synonym = original.Synonyms.FirstOrDefault(s => TermNames.Normalize(s) == synonym) ?? synonym;
            if (synonym.Length == 0
                || string.Equals(synonym, name, StringComparison.OrdinalIgnoreCase)
                || synonyms.Contains(synonym, StringComparer.OrdinalIgnoreCase)) continue;
            error ??= TermNames.Validate(synonym, "synonym");
            synonyms.Add(synonym);
        }

        error = TermNames.Validate(name, Node?.Kind == TermNodeKind.Term ? "default label" : "name")
            ?? TermNames.ValidateDescription(description)
            ?? error;

        var custom = ReadRows(_customProps, "custom property", ref error);
        var local = ReadRows(_localProps, "local property", ref error);

        var edited = original with
        {
            Name = name,
            Description = description,
            Synonyms = synonyms,
            Owner = (OwnerBox.Text ?? "").Trim(),
            Contact = (ContactBox.Text ?? "").Trim(),
            IsOpenForTermCreation = OpenBox.IsOn,
            IsAvailableForTagging = TaggingBox.IsOn,
            CustomProperties = custom,
            LocalCustomProperties = local,
        };
        return error == null ? edited : null;
    }

    /// <summary>Shows a problem that stops the save (validation, or a precondition the window checks).</summary>
    public void ShowValidationError(string message)
    {
        ValidationBar.Message = message;
        ValidationBar.IsOpen = true;
        UpdateHintPanel();
    }

    private static List<PropertyEntry> ReadRows(IEnumerable<PropertyRow> rows, string what, ref string? error)
    {
        var entries = new List<PropertyEntry>();
        foreach (var row in rows)
        {
            var key = (row.Key ?? "").Trim();
            var value = row.Value ?? "";
            if (key.Length == 0)
            {
                if (value.Trim().Length > 0) error ??= $"A {what} has a value but no name.";
                continue; // blank row
            }
            if (entries.Any(e => e.Key == key))
            {
                error ??= $"There are two {what} rows named “{key}”.";
                continue;
            }
            entries.Add(new PropertyEntry(key, value));
        }
        return entries;
    }

    // ---- Events --------------------------------------------------------------------------------------------

    private void Field_Changed(object sender, RoutedEventArgs e) => UpdateDirty();

    private void Rows_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        foreach (PropertyRow row in e.OldItems ?? Array.Empty<PropertyRow>()) row.PropertyChanged -= Row_PropertyChanged;
        foreach (PropertyRow row in e.NewItems ?? Array.Empty<PropertyRow>()) row.PropertyChanged += Row_PropertyChanged;
        UpdateDirty();
    }

    private void Row_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(PropertyRow.IsEditable)) UpdateDirty();
    }

    // ---- Synonyms ------------------------------------------------------------------------------------------

    /// <summary>
    /// Several synonyms can be typed at once, separated by semicolons. Semicolons aren't allowed in labels, so
    /// nothing is lost; commas are allowed (e.g. "Smith, John"), so they don't separate.
    /// </summary>
    private static IEnumerable<string> SplitSynonyms(string? text) =>
        (text ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    private void SynonymInput_TextChanged(object sender, TextChangedEventArgs e)
    {
        AddSynonymButton.IsEnabled = !_readOnly && !_busy && !string.IsNullOrWhiteSpace(SynonymInput.Text);
        UpdateDirty();
    }

    private void SynonymInput_KeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (e.Key != VirtualKey.Enter) return;
        e.Handled = true;
        AddSynonyms();
    }

    private void AddSynonym_Click(object sender, RoutedEventArgs e) => AddSynonyms();

    /// <summary>
    /// Adds what's typed as chips: each part tidied as SharePoint stores it, skipping blanks, ones already there
    /// and the default label. Parts that break the label rules stay in the box with the reason shown below it.
    /// </summary>
    private void AddSynonyms()
    {
        if (_readOnly || _busy) return;
        var defaultLabel = TermNames.Normalize(NameBox.Text);
        var rejected = new List<string>();
        string? error = null;
        foreach (var part in SplitSynonyms(SynonymInput.Text))
        {
            var synonym = TermNames.Normalize(part);
            if (synonym.Length == 0
                || string.Equals(synonym, defaultLabel, StringComparison.OrdinalIgnoreCase)
                || _synonyms.Any(s => string.Equals(TermNames.Normalize(s), synonym, StringComparison.OrdinalIgnoreCase))) continue;
            if (TermNames.Validate(synonym, "synonym") is { } problem)
            {
                error ??= problem;
                rejected.Add(part);
                continue;
            }
            _synonyms.Add(synonym);
        }

        RebuildSynonymChips();
        SynonymError.Text = error ?? "";
        SynonymError.Visibility = Vis(error != null);
        SynonymInput.Text = string.Join("; ", rejected);
        SynonymInput.Focus(FocusState.Programmatic);
        UpdateDirty();
    }

    private void RemoveSynonym(string synonym)
    {
        _synonyms.Remove(synonym);
        RebuildSynonymChips();
        SynonymInput.Focus(FocusState.Programmatic);
        UpdateDirty();
    }

    /// <summary>One chip per synonym: the label (shortened with "…" if wider than the panel) and a remove button.</summary>
    private void RebuildSynonymChips()
    {
        SynonymChips.Children.Clear();
        _synonymRemoveButtons.Clear();
        bool editable = !_readOnly && !_busy;
        foreach (var synonym in _synonyms)
        {
            var remove = new Button
            {
                Content = new FontIcon { Glyph = FluentIcons.Dismiss, FontFamily = FluentIcons.RegularFontFamily, FontSize = 12 },
                Style = (Style)Resources["SynonymRemoveButtonStyle"],
                IsEnabled = editable,
            };
            AutomationProperties.SetName(remove, $"Remove synonym {synonym}");
            ToolTipService.SetToolTip(remove, "Remove");
            remove.Click += (_, _) => RemoveSynonym(synonym);
            _synonymRemoveButtons.Add(remove);

            var label = new TextBlock { Text = synonym, VerticalAlignment = VerticalAlignment.Center, TextTrimming = TextTrimming.CharacterEllipsis };
            ToolTipService.SetToolTip(label, synonym);
            var content = new Grid();
            content.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            content.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            Grid.SetColumn(remove, 1);
            content.Children.Add(label);
            content.Children.Add(remove);

            SynonymChips.Children.Add(new Border { Child = content, Style = (Style)Resources["SynonymChipStyle"] });
        }
    }

    private void AddCustomProperty_Click(object sender, RoutedEventArgs e) => _customProps.Add(new PropertyRow());

    private void AddLocalProperty_Click(object sender, RoutedEventArgs e) => _localProps.Add(new PropertyRow());

    private void RemoveProperty_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is not PropertyRow row) return;
        if (!_customProps.Remove(row)) _localProps.Remove(row);
    }

    /// <summary>The hint and validation bars take space (and their gap) only while one is open.</summary>
    private void UpdateHintPanel() => HintPanel.Visibility = Vis(HintBar.IsOpen || ValidationBar.IsOpen);

    private void ValidationBar_Closed(InfoBar sender, InfoBarClosedEventArgs args) => UpdateHintPanel();

    private void CopyId_Click(object sender, RoutedEventArgs e) => CopyText(IdText.Text);

    private void CopyNameId_Click(object sender, RoutedEventArgs e) => CopyText(NameIdText.Text);

    private static void CopyText(string text)
    {
        var package = new Windows.ApplicationModel.DataTransfer.DataPackage();
        package.SetText(text);
        Windows.ApplicationModel.DataTransfer.Clipboard.SetContent(package);
    }

    /// <summary>
    /// Sets the header to the item's path as a breadcrumb, "Group › Term set › … › <paramref name="name"/>": the
    /// ancestors are buttons that select them, the item itself is bold. With no item, hides the header and shows
    /// the centered placeholder instead.
    /// </summary>
    private void SetTitle(string name, TermNode? node)
    {
        TitlePanel.Visibility = Vis(node != null);
        EmptyState.Visibility = Vis(node == null);
        Breadcrumb.Children.Clear();
        if (node == null) return;

        var ancestors = new List<TermNode>();
        for (var parent = node.Parent; parent != null; parent = parent.Parent) ancestors.Insert(0, parent);

        // BreadcrumbPanel's children: the overflow "…" (a menu of the ancestors it hides when space is short),
        // the ancestors, then the item.
        Breadcrumb.Children.Add(BreadcrumbItem("…", "Show the rest of the path", button => ShowHiddenAncestors(button, ancestors)));
        foreach (var ancestor in ancestors)
            Breadcrumb.Children.Add(BreadcrumbItem(ancestor.Name, ancestor.Name, _ => PathItemInvoked?.Invoke(this, ancestor)));
        var current = new TextBlock { Text = name, Style = (Style)Resources["BreadcrumbCurrentStyle"] };
        ToolTipService.SetToolTip(current, name);
        Breadcrumb.Children.Add(current);
    }

    /// <summary>A breadcrumb button (its text shortened with "…" when the panel gives it less room) and the chevron after it.</summary>
    private Grid BreadcrumbItem(string text, string toolTip, Action<Button> click)
    {
        var button = new Button
        {
            Content = new TextBlock { Text = text, TextTrimming = TextTrimming.CharacterEllipsis, TextWrapping = TextWrapping.NoWrap },
            Style = (Style)Resources["BreadcrumbButtonStyle"],
        };
        AutomationProperties.SetName(button, text);
        ToolTipService.SetToolTip(button, toolTip);
        button.Click += (_, _) => click(button);

        var chevron = new FontIcon { Glyph = FluentIcons.ChevronRight, Style = (Style)Resources["BreadcrumbChevronStyle"] };
        var item = new Grid();
        item.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        item.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        Grid.SetColumn(chevron, 1);
        item.Children.Add(button);
        item.Children.Add(chevron);
        return item;
    }

    /// <summary>The overflow menu: the ancestors the breadcrumb is hiding for lack of room.</summary>
    private void ShowHiddenAncestors(Button button, List<TermNode> ancestors)
    {
        var menu = new MenuFlyout();
        foreach (var ancestor in ancestors.Take(Breadcrumb.HiddenCount))
        {
            var item = new MenuFlyoutItem { Text = ancestor.Name };
            item.Click += (_, _) => PathItemInvoked?.Invoke(this, ancestor);
            menu.Items.Add(item);
        }
        if (menu.Items.Count > 0) menu.ShowAt(button);
    }

    private void Discard_Click(object sender, RoutedEventArgs e)
    {
        if (Node != null && _original != null) Show(Node, _original);
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        if (!IsDirty || _busy || _readOnly) return;
        SaveRequested?.Invoke(this, EventArgs.Empty);
    }

    // ---- Helpers -------------------------------------------------------------------------------------------

    private void UpdateDirty()
    {
        if (_loadingForm || _original == null) return;
        var edited = GetEdited(out _);
        // Invalid input still counts as a change, so Save stays enabled and reports what to fix.
        SetDirty(edited == null || !edited.SameEditableValues(_original));
    }

    private void SetDirty(bool dirty)
    {
        bool changed = IsDirty != dirty;
        IsDirty = dirty;
        DirtyText.Visibility = Vis(dirty);
        UpdateEnabled();
        if (changed) DirtyChanged?.Invoke(this, EventArgs.Empty);
    }

    private void UpdateEnabled()
    {
        bool editable = !_readOnly && !_busy;
        foreach (var control in new Control[] { NameBox, DescriptionBox, SynonymInput, OwnerBox, ContactBox, OpenBox, TaggingBox, AddCustomPropButton, AddLocalPropButton })
            control.IsEnabled = editable;
        AddSynonymButton.IsEnabled = editable && !string.IsNullOrWhiteSpace(SynonymInput.Text);
        foreach (var button in _synonymRemoveButtons) button.IsEnabled = editable;
        foreach (var row in _customProps.Concat(_localProps)) row.IsEditable = editable;
        SaveButton.IsEnabled = editable && IsDirty;
        DiscardButton.IsEnabled = editable && IsDirty;
    }

    private void FillRows(ObservableCollection<PropertyRow> rows, IEnumerable<PropertyEntry> entries)
    {
        rows.Clear();
        foreach (var entry in entries) rows.Add(new PropertyRow { Key = entry.Key, Value = entry.Value });
    }

    private static string KeepOriginal(string value, string original, Func<string, string> normalize) =>
        normalize(original) == value ? original : value;

    private static string NormalizeDescription(string? text) =>
        (text ?? "").Trim().Replace("\r\n", "\n").Replace('\r', '\n');

    private static Visibility Vis(bool visible) => visible ? Visibility.Visible : Visibility.Collapsed;

    private static string FormatDate(DateTime date) =>
        date == default ? "—" : (date.Kind == DateTimeKind.Utc ? date.ToLocalTime() : date).ToString("g", CultureInfo.CurrentCulture);

    private static string LanguageName(int lcid) => Services.TermStoreService.LanguageName(lcid);

    /// <summary>The icon beside a "Member of" relation: source, pinned or reused copy.</summary>
    public static string RelationGlyph(string relation) => relation switch
    {
        "Source" => FluentIcons.CheckmarkCircle,
        "Pinned" => FluentIcons.Pin,
        _ => FluentIcons.Link,
    };

    private static string StatusOf(NodeDetails d)
    {
        var flags = new List<string>();
        if (d.IsDeprecated) flags.Add("Deprecated");
        if (d.IsPinned) flags.Add(d.IsSourceTerm ? "Pinned (source)" : "Pinned copy");
        else if (d.IsReused) flags.Add(d.IsSourceTerm ? "Reused (source term)" : "Reused");
        flags.Add(d.TermsCount == 1 ? "1 child term" : $"{d.TermsCount} child terms");
        return string.Join(" · ", flags);
    }
}
