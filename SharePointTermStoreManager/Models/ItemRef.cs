using System.Text.Json.Serialization;
using SharePointTermStoreManager.Helpers;

namespace SharePointTermStoreManager.Models;

/// <summary>
/// Identifies a term group, term set or term without needing it in the tree, e.g. a search result in the target
/// picker. A term is identified by its ID plus the term set it's in (<see cref="TermSetId"/>), because a reused
/// term has the same ID in every term set it's reused in.
/// </summary>
public sealed record ItemRef(TermNodeKind Kind, Guid Id, Guid? TermSetId, string Name, string Path)
{
    public bool IsDeprecated { get; init; }
    public bool IsReused { get; init; }
    public bool IsPinned { get; init; }
    public bool IsSourceTerm { get; init; }
    public int ChildTermCount { get; init; }
    public bool IsSystemGroup { get; init; }

    /// <summary>For search results: the group the term set is in, and the default labels of the term's parents, top first.</summary>
    public Guid? GroupId { get; init; }
    public IReadOnlyList<string> ParentLabels { get; init; } = [];

    [JsonIgnore]
    public bool IsPinnedCopy => Kind == TermNodeKind.Term && IsPinned && !IsSourceTerm;

    // For the picker's result list.
    [JsonIgnore]
    public string Glyph => TermIcons.Outline(Kind, IsDeprecated, IsReused, IsPinned);

    [JsonIgnore]
    public double TextOpacity => IsDeprecated ? 0.55 : 1;

    [JsonIgnore]
    public string Badges => string.Join(" · ", new[]
    {
        IsDeprecated ? "Deprecated" : null,
        IsPinnedCopy ? "Pinned copy" : IsReused ? "Reused" : null,
    }.Where(b => b != null));
}
