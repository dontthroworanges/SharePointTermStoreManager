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

    /// <summary>This node and everything loaded below it (shown in the tree or not; never the "Show more" row).</summary>
    public IEnumerable<TermNode> SelfAndDescendants()
    {
        yield return this;
        foreach (var child in LoadedChildren)
        foreach (var node in child.SelfAndDescendants())
            yield return node;
    }

    // ---- Paging: the tree shows a page of children at a time ----------------------------------------------------

    /// <summary>
    /// True for the "Show more" row at the end of a partly shown list of children: not a real item, so it has no
    /// commands and isn't part of <see cref="LoadedChildren"/>.
    /// </summary>
    public bool IsMoreItem { get; set; }

    /// <summary>Children that are loaded but not shown in the tree yet, in order, after the shown ones.</summary>
    public List<TermNode> HiddenChildren { get; } = new();

    /// <summary>Every loaded child, in order: the shown ones, then the hidden ones (never the "Show more" row).</summary>
    public IEnumerable<TermNode> LoadedChildren => Children.Where(c => !c.IsMoreItem).Concat(HiddenChildren);

    /// <summary>The page size last used for this node's children; 0 shows them all.</summary>
    private int _pageSize;

    /// <summary>Replaces the children, showing the first <paramref name="pageSize"/> (0: all) and a "Show more" row for the rest.</summary>
    public void SetChildren(IEnumerable<TermNode> children, int pageSize)
    {
        _pageSize = Math.Max(0, pageSize);
        var list = children.ToList();
        int shown = _pageSize == 0 ? list.Count : Math.Min(_pageSize, list.Count);
        Children.Clear();
        HiddenChildren.Clear();
        foreach (var child in list.Take(shown)) Children.Add(child);
        HiddenChildren.AddRange(list.Skip(shown));
        UpdateMoreItem();
    }

    /// <summary>Removes every child, shown or hidden.</summary>
    public void ClearChildren() => SetChildren([], _pageSize);

    /// <summary>
    /// Shows the next page of hidden children, or, with <paramref name="until"/>, as many pages as it takes to show
    /// that child.
    /// </summary>
    public void ShowMore(TermNode? until = null)
    {
        if (HiddenChildren.Count == 0) return;
        int count = _pageSize == 0 ? HiddenChildren.Count : Math.Min(_pageSize, HiddenChildren.Count);
        if (until != null && HiddenChildren.IndexOf(until) is var index and >= 0 && index >= count)
            count = _pageSize == 0 ? HiddenChildren.Count : Math.Min(HiddenChildren.Count, (index / _pageSize + 1) * _pageSize);

        RemoveMoreItem();
        foreach (var child in HiddenChildren.Take(count)) Children.Add(child);
        HiddenChildren.RemoveRange(0, count);
        UpdateMoreItem();
    }

    /// <summary>
    /// Inserts a child at <paramref name="index"/> in <see cref="LoadedChildren"/> order, showing it even when that
    /// position is among the hidden ones (so it can be selected straight away).
    /// </summary>
    public void InsertChild(int index, TermNode child)
    {
        int shown = Children.Count(c => !c.IsMoreItem);
        if (index <= shown)
        {
            Children.Insert(index, child);
            return;
        }
        HiddenChildren.Insert(Math.Min(index - shown, HiddenChildren.Count), child);
        ShowMore(child);
    }

    /// <summary>Removes a child, shown or hidden. False when it isn't one.</summary>
    public bool RemoveChild(TermNode child)
    {
        bool removed = Children.Remove(child) || HiddenChildren.Remove(child);
        if (removed) UpdateMoreItem();
        return removed;
    }

    private void RemoveMoreItem()
    {
        if (Children.LastOrDefault() is { IsMoreItem: true } more) Children.Remove(more);
    }

    /// <summary>Puts the "Show more" row at the end while there are hidden children, with the count left.</summary>
    private void UpdateMoreItem()
    {
        RemoveMoreItem();
        if (HiddenChildren.Count == 0) return;
        Children.Add(new TermNode
        {
            IsMoreItem = true,
            Kind = HiddenChildren[0].Kind,
            Parent = this,
            Name = $"Show more ({HiddenChildren.Count:N0} not shown)",
        });
    }

    /// <summary>Tree row parts: the item's icon, or the "…" icon of the "Show more" row.</summary>
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
