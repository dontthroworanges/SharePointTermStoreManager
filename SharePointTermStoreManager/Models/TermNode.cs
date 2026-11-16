using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.UI.Xaml;
using SharePointTermStoreManager.Helpers;

namespace SharePointTermStoreManager.Models;

public enum TermNodeKind
{
    Group,
    TermSet,
    Term,
}

/// <summary>
/// One node in the term store tree: a term group, a term set or a term. Children are loaded when the node is
/// first expanded (<see cref="HasUnrealizedChildren"/>), so large term sets don't have to be read up front.
/// </summary>
public sealed partial class TermNode : ObservableObject
{
    public TermNodeKind Kind { get; set; }
    public Guid Id { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Path))]
    private string _name = "";

    [ObservableProperty]
    private string _description = "";

    /// <summary>The parent node; null for term groups.</summary>
    public TermNode? Parent { get; set; }

    /// <summary>
    /// True when this term set or term orders its children by a custom sort order. New child terms are then added
    /// at the end, as SharePoint does, rather than alphabetically.
    /// </summary>
    public bool HasCustomSortOrder { get; set; }

    /// <summary>The term set this node is in (itself for a term set); null for groups.</summary>
    public TermNode? TermSet => Kind switch
    {
        TermNodeKind.Group => null,
        TermNodeKind.TermSet => this,
        _ => Parent?.TermSet,
    };

    /// <summary>A pinned copy of a reused term: SharePoint only allows changes at the source term.</summary>
    public bool IsPinnedCopy => Kind == TermNodeKind.Term && IsPinned && !IsSourceTerm;

    /// <summary>This node and everything loaded below it (shown in the tree or not; never the "Show earlier" / "Show more" rows).</summary>
    public IEnumerable<TermNode> SelfAndDescendants()
    {
        yield return this;
        foreach (var child in LoadedChildren)
        foreach (var node in child.SelfAndDescendants())
            yield return node;
    }

    // ---- Paging: the tree shows a window of children, a page at a time ------------------------------------------
    //
    // The shown children are a contiguous window of all the loaded ones. Children before the window are behind a
    // "Show earlier" row at the top, children after it behind a "Show more" row at the end. Going to a hidden child
    // (e.g. a search result near the end of a large term set) moves the window to it rather than showing
    // everything above it, so the tree stays short and focused on that item.

    /// <summary>
    /// True for the "Show earlier" / "Show more" rows of a partly shown list of children: not a real item, so it
    /// has no commands and isn't part of <see cref="LoadedChildren"/>.
    /// </summary>
    public bool IsMoreItem { get; set; }

    /// <summary>True for the "Show earlier" row at the top of a partly shown list (with <see cref="IsMoreItem"/>).</summary>
    public bool IsEarlierItem { get; set; }

    /// <summary>Loaded children not shown in the tree, in order: those before the shown ones and those after.</summary>
    private readonly List<TermNode> _hiddenBefore = new();
    private readonly List<TermNode> _hiddenAfter = new();

    private IEnumerable<TermNode> ShownChildren => Children.Where(c => !c.IsMoreItem);

    /// <summary>Every loaded child, in order, shown in the tree or not (never the "Show earlier" / "Show more" rows).</summary>
    public IEnumerable<TermNode> LoadedChildren => _hiddenBefore.Concat(ShownChildren).Concat(_hiddenAfter);

    /// <summary>True when <paramref name="child"/> is loaded but not shown in the tree.</summary>
    public bool IsHidden(TermNode child) => _hiddenBefore.Contains(child) || _hiddenAfter.Contains(child);

    /// <summary>The page size last used for this node's children; 0 shows them all.</summary>
    private int _pageSize;

    /// <summary>Replaces the children, showing the first <paramref name="pageSize"/> (0: all) and a "Show more" row for the rest.</summary>
    public void SetChildren(IEnumerable<TermNode> children, int pageSize)
    {
        _pageSize = Math.Max(0, pageSize);
        ShowWindow(children.ToList(), 0);
    }

    /// <summary>Removes every child, shown or hidden.</summary>
    public void ClearChildren() => SetChildren([], _pageSize);

    /// <summary>Shows the next page of the children after the shown ones.</summary>
    public void ShowMore()
    {
        if (_hiddenAfter.Count == 0) return;
        int count = PageCount(_hiddenAfter.Count);
        RemoveMoreItems();
        foreach (var child in _hiddenAfter.Take(count)) Children.Add(child);
        _hiddenAfter.RemoveRange(0, count);
        UpdateMoreItems();
    }

    /// <summary>Shows the previous page of the children before the shown ones.</summary>
    public void ShowEarlier()
    {
        if (_hiddenBefore.Count == 0) return;
        int count = PageCount(_hiddenBefore.Count);
        int start = _hiddenBefore.Count - count;
        RemoveMoreItems();
        for (int i = 0; i < count; i++) Children.Insert(i, _hiddenBefore[start + i]);
        _hiddenBefore.RemoveRange(start, count);
        UpdateMoreItems();
    }

    /// <summary>
    /// Shows a hidden child by moving the shown window to a page around it (the rest stay behind "Show earlier" and
    /// "Show more"), rather than showing everything between. Does nothing when it's already shown.
    /// </summary>
    public void ShowPageOf(TermNode child)
    {
        if (!IsHidden(child)) return;
        var all = LoadedChildren.ToList();
        int index = all.IndexOf(child);
        int count = PageCount(all.Count);
        // The child in the middle of the page, as far as the ends of the list allow.
        int start = Math.Clamp(index - count / 2, 0, all.Count - count);
        ShowWindow(all, start);
    }

    /// <summary>
    /// Inserts a child at <paramref name="index"/> in <see cref="LoadedChildren"/> order, showing it even when that
    /// position is among the hidden ones (so it can be selected straight away).
    /// </summary>
    public void InsertChild(int index, TermNode child)
    {
        int before = _hiddenBefore.Count;
        int shown = ShownChildren.Count();
        if (index < before)
        {
            _hiddenBefore.Insert(index, child);
            ShowPageOf(child);
        }
        else if (index <= before + shown)
        {
            // After the "Show earlier" row when there is one; before the "Show more" row.
            Children.Insert(index - before + (before > 0 ? 1 : 0), child);
        }
        else
        {
            _hiddenAfter.Insert(Math.Min(index - before - shown, _hiddenAfter.Count), child);
            ShowPageOf(child);
        }
    }

    /// <summary>Removes a child, shown or hidden. False when it isn't one.</summary>
    public bool RemoveChild(TermNode child)
    {
        bool removed = Children.Remove(child) || _hiddenBefore.Remove(child) || _hiddenAfter.Remove(child);
        if (removed) UpdateMoreItems();
        return removed;
    }

    /// <summary>How many children a page holds out of <paramref name="available"/>.</summary>
    private int PageCount(int available) => _pageSize == 0 ? available : Math.Min(_pageSize, available);

    /// <summary>Shows a page of <paramref name="all"/> starting at <paramref name="start"/>, hiding the rest.</summary>
    private void ShowWindow(List<TermNode> all, int start)
    {
        int count = PageCount(all.Count - start);
        Children.Clear();
        _hiddenBefore.Clear();
        _hiddenAfter.Clear();
        _hiddenBefore.AddRange(all.Take(start));
        foreach (var child in all.Skip(start).Take(count)) Children.Add(child);
        _hiddenAfter.AddRange(all.Skip(start + count));
        UpdateMoreItems();
    }

    private void RemoveMoreItems()
    {
        if (Children.FirstOrDefault() is { IsMoreItem: true } earlier) Children.Remove(earlier);
        if (Children.LastOrDefault() is { IsMoreItem: true } more) Children.Remove(more);
    }

    /// <summary>
    /// Puts the "Show earlier" row at the top and the "Show more" row at the end while there are hidden children
    /// on that side, with the count left.
    /// </summary>
    private void UpdateMoreItems()
    {
        RemoveMoreItems();
        if (_hiddenBefore.Count > 0)
        {
            Children.Insert(0, new TermNode
            {
                IsMoreItem = true,
                IsEarlierItem = true,
                Kind = _hiddenBefore[^1].Kind,
                Parent = this,
                Name = $"Show earlier ({_hiddenBefore.Count:N0} not shown)",
            });
        }
        if (_hiddenAfter.Count > 0)
        {
            Children.Add(new TermNode
            {
                IsMoreItem = true,
                Kind = _hiddenAfter[0].Kind,
                Parent = this,
                Name = $"Show more ({_hiddenAfter.Count:N0} not shown)",
            });
        }
    }

    /// <summary>Tree row parts: the item's icon, or the "…" icon of the "Show earlier" / "Show more" rows.</summary>
    public Visibility ItemIconVisibility => IsMoreItem ? Visibility.Collapsed : Visibility.Visible;
    public Visibility MoreIconVisibility => IsMoreItem ? Visibility.Visible : Visibility.Collapsed;

    // Group flags.
    public bool IsSystemGroup { get; set; }
    public bool IsSiteCollectionGroup { get; set; }

    // Term flags. Observable so the tree's icon and dimming follow deprecate / reuse / pin.
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(TextOpacity), nameof(Flags), nameof(Glyph))]
    private bool _isDeprecated;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Glyph), nameof(Flags))]
    private bool _isReused;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Glyph), nameof(Flags))]
    private bool _isPinned;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Flags))]
    private bool _isSourceTerm;

    public int ChildTermCount { get; set; }

    /// <summary>A reference to this item for operations and pickers.</summary>
    public ItemRef ToRef() => new(Kind, Id, TermSet?.Id, Name, Path)
    {
        IsDeprecated = IsDeprecated,
        IsReused = IsReused,
        IsPinned = IsPinned,
        IsSourceTerm = IsSourceTerm,
        ChildTermCount = ChildTermCount,
        IsSystemGroup = IsSystemGroup,
    };

    /// <summary>True when this node is the item <paramref name="item"/> refers to (the same instance, for terms).</summary>
    public bool Matches(ItemRef item) =>
        Kind == item.Kind && Id == item.Id && (Kind != TermNodeKind.Term || TermSet?.Id == item.TermSetId);

    public ObservableCollection<TermNode> Children { get; } = new();

    [ObservableProperty]
    private bool _hasUnrealizedChildren;

    /// <summary>Bound two-way to the tree item, so code can open a path (e.g. to reveal a search result).</summary>
    [ObservableProperty]
    private bool _isExpanded;

    /// <summary>True while this node's children are being read from SharePoint.</summary>
    [ObservableProperty]
    private bool _isLoading;

    public string Glyph => TermIcons.Outline(Kind, IsDeprecated, IsReused, IsPinned);

    /// <summary>Deprecated terms are dimmed, as in the SharePoint term store.</summary>
    public double TextOpacity => IsDeprecated ? 0.55 : 1;

    public string KindLabel => Kind switch
    {
        TermNodeKind.Group => IsSiteCollectionGroup ? "Site collection term group" : IsSystemGroup ? "System term group" : "Term group",
        TermNodeKind.TermSet => "Term set",
        _ => "Term",
    };

    /// <summary>Lower-case noun for messages: "term group", "term set", "term".</summary>
    public static string Noun(TermNodeKind kind) => kind switch
    {
        TermNodeKind.Group => "term group",
        TermNodeKind.TermSet => "term set",
        _ => "term",
    };

    /// <summary>"Group › Term set › Parent › Term".</summary>
    public string Path => Parent == null ? Name : $"{Parent.Path} › {Name}";

    /// <summary>Comma-separated term flags for the details pane, e.g. "Deprecated, Reused".</summary>
    public string Flags
    {
        get
        {
            var flags = new List<string>();
            if (IsDeprecated) flags.Add("Deprecated");
            if (IsPinned) flags.Add("Pinned");
            else if (IsReused) flags.Add("Reused");
            if (IsReused && IsSourceTerm) flags.Add("Source term");
            return flags.Count == 0 ? "None" : string.Join(", ", flags);
        }
    }
}
