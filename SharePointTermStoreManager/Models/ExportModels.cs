using System.Text.Json.Serialization;

namespace SharePointTermStoreManager.Models;

/// <summary>The file formats the Export page writes.</summary>
public enum ExportFormat
{
    Json,
    Excel,
    Csv,
    Html,
}

/// <summary>Progress of reading the term store for an export.</summary>
public sealed record ExportProgress(int Done, int Total, string Text);

/// <summary>
/// The term store as exported: its term groups, their term sets and every term, nested as in the tree, with
/// each item's details (descriptions, synonyms, custom properties, flags, dates and IDs). Serialized as is for
/// the JSON export; flattened into <see cref="ExportRow"/>s for Excel and CSV.
/// </summary>
public sealed class ExportStore
{
    public Guid Id { get; init; }
    public string Name { get; init; } = "";
    public string Site { get; init; } = "";
    public DateTimeOffset Exported { get; init; }

    /// <summary>LCID of the term store's default language, which names, descriptions and synonyms are in.</summary>
    public int DefaultLanguage { get; init; }

    public List<ExportGroup> Groups { get; init; } = [];

    [JsonIgnore]
    public int TermSetCount => Groups.Sum(g => g.TermSets.Count);

    [JsonIgnore]
    public int TermCount => Groups.Sum(g => g.TermSets.Sum(s => s.Terms.Sum(t => t.SelfAndDescendantCount)));
}

public sealed class ExportGroup
{
    public Guid Id { get; init; }
    public string Name { get; init; } = "";
    public string Description { get; init; } = "";
    public bool IsSystemGroup { get; init; }
    public bool IsSiteCollectionGroup { get; init; }
    public DateTime Created { get; init; }
    public DateTime LastModified { get; init; }
    public List<ExportTermSet> TermSets { get; init; } = [];
}

public sealed class ExportTermSet
{
    public Guid Id { get; init; }
    public string Name { get; init; } = "";
    public string Description { get; init; } = "";
    public string Owner { get; init; } = "";
    public string Contact { get; init; } = "";
    public bool IsOpenForTermCreation { get; init; }
    public bool IsAvailableForTagging { get; init; }
    public bool HasCustomSortOrder { get; init; }
    public Dictionary<string, string> CustomProperties { get; init; } = [];
    public DateTime Created { get; init; }
    public DateTime LastModified { get; init; }
    public List<ExportTerm> Terms { get; set; } = [];
}

public sealed class ExportTerm
{
    public Guid Id { get; init; }

    /// <summary>The default label in the term store's default language.</summary>
    public string Name { get; init; } = "";

    public string Description { get; init; } = "";

    /// <summary>The other labels in the default language.</summary>
    public List<string> Synonyms { get; init; } = [];

    /// <summary>Labels in the term store's other languages; null (left out) when there are none.</summary>
    public List<ExportLanguageLabels>? OtherLanguages { get; init; }

    public bool IsDeprecated { get; init; }
    public bool IsAvailableForTagging { get; init; }
    public bool IsReused { get; init; }
    public bool IsPinned { get; init; }
    public bool IsSourceTerm { get; init; }
    public bool HasCustomSortOrder { get; init; }
    public Dictionary<string, string> CustomProperties { get; init; } = [];
    public Dictionary<string, string> LocalCustomProperties { get; init; } = [];
    public DateTime Created { get; init; }
    public DateTime LastModified { get; init; }
    public List<ExportTerm> Terms { get; set; } = [];

    /// <summary>The parent's custom sort order, used while building the tree.</summary>
    [JsonIgnore]
    public string? CustomSortOrder { get; init; }

    [JsonIgnore]
    public int SelfAndDescendantCount => 1 + Terms.Sum(t => t.SelfAndDescendantCount);
}

/// <summary>A term's default label and synonyms in one of the term store's other languages.</summary>
public sealed record ExportLanguageLabels(int Lcid, string Language, string Name, List<string> Synonyms);

/// <summary>One item (term group, term set or term) as a table row, for the Excel and CSV exports.</summary>
public sealed class ExportRow
{
    public string Type { get; init; } = "";
    public string TermGroup { get; init; } = "";
    public string TermSet { get; init; } = "";

    /// <summary>A term's depth in its term set (1 for top-level terms); null for groups and term sets.</summary>
    public int? Level { get; init; }

    public string Name { get; init; } = "";
    public string Path { get; init; } = "";
    public string Description { get; init; } = "";
    public string Synonyms { get; init; } = "";
    public string OtherLanguages { get; init; } = "";
    public bool? AvailableForTagging { get; init; }
    public bool? OpenForTermCreation { get; init; }
    public bool? Deprecated { get; init; }
    public bool? Reused { get; init; }
    public bool? Pinned { get; init; }
    public bool? SourceTerm { get; init; }
    public string Owner { get; init; } = "";
    public string Contact { get; init; } = "";
    public string CustomProperties { get; init; } = "";
    public string LocalCustomProperties { get; init; } = "";
    public int? ChildTerms { get; init; }
    public Guid Id { get; init; }
    public Guid? ParentId { get; init; }
    public Guid? TermSetId { get; init; }
    public Guid TermGroupId { get; init; }
    public DateTime Created { get; init; }
    public DateTime LastModified { get; init; }
}
