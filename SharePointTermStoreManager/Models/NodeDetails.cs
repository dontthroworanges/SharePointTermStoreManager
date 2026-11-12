namespace SharePointTermStoreManager.Models;

/// <summary>A custom property (shared or local) of a term set or term.</summary>
public sealed record PropertyEntry(string Key, string Value);

/// <summary>A term set a term is in, and how: "Source", "Reused" or "Pinned" (the term store's "Member of").</summary>
public sealed record TermSetMembership(string TermSetName, string Description, string Relation);

/// <summary>A term's labels in a language other than the term store's default language (shown read-only).</summary>
public sealed record LanguageLabels(int Lcid, string Language, string Label, string Synonyms);

/// <summary>
/// The editable properties of a term group, term set or term, read fresh from SharePoint when the item is
/// selected. Labels and descriptions are in the term store's default language (<see cref="Language"/>); labels in
/// other languages are left untouched when saving.
/// </summary>
public sealed record NodeDetails
{
    public TermNodeKind Kind { get; init; }
    public Guid Id { get; init; }

    /// <summary>LCID of the language <see cref="Name"/>, <see cref="Description"/> and <see cref="Synonyms"/> are in.</summary>
    public int Language { get; init; }

    /// <summary>Name of a group or term set; default label of a term.</summary>
    public string Name { get; init; } = "";
    public string Description { get; init; } = "";

    public DateTime Created { get; init; }

    /// <summary>Checked before saving, so changes made elsewhere since the item was opened aren't overwritten.</summary>
    public DateTime LastModified { get; init; }

    // Group.
    public bool IsSystemGroup { get; init; }

    // Term set.
    public string Owner { get; init; } = "";
    public string Contact { get; init; } = "";
    public bool IsOpenForTermCreation { get; init; }

    // Term set and term.
    public bool IsAvailableForTagging { get; init; }

    /// <summary>True when the children are in a custom order rather than alphabetical (read-only here).</summary>
    public bool HasCustomSortOrder { get; init; }
    public IReadOnlyList<PropertyEntry> CustomProperties { get; init; } = [];

    // Term.
    /// <summary>Labels in <see cref="Language"/> other than the default label.</summary>
    public IReadOnlyList<string> Synonyms { get; init; } = [];
    public IReadOnlyList<PropertyEntry> LocalCustomProperties { get; init; } = [];

    /// <summary>The term sets this term is in, this one first.</summary>
    public IReadOnlyList<TermSetMembership> MemberOf { get; init; } = [];

    /// <summary>Labels in the term store's other languages, for reference.</summary>
    public IReadOnlyList<LanguageLabels> OtherLanguages { get; init; } = [];
    public bool IsDeprecated { get; init; }
    public bool IsReused { get; init; }
    public bool IsPinned { get; init; }
    public bool IsSourceTerm { get; init; }
    public int TermsCount { get; init; }

    /// <summary>True when the editable values of the two match (ignoring order of synonyms and properties).</summary>
    public bool SameEditableValues(NodeDetails other) =>
        Name == other.Name
        && Description == other.Description
        && Owner == other.Owner
        && Contact == other.Contact
        && IsOpenForTermCreation == other.IsOpenForTermCreation
        && IsAvailableForTagging == other.IsAvailableForTagging
        && SameSet(Synonyms, other.Synonyms)
        && SameProperties(CustomProperties, other.CustomProperties)
        && SameProperties(LocalCustomProperties, other.LocalCustomProperties);

    /// <summary>
    /// What a save changes, for the operations log: each changed field with its old and new value; for synonyms
    /// and custom properties, what was added, removed or changed. Fields that didn't change are left out.
    /// </summary>
    public Dictionary<string, object> ChangesTo(NodeDetails edited)
    {
        var changes = new Dictionary<string, object>();
        void Value<T>(string field, T before, T after)
        {
            if (!EqualityComparer<T>.Default.Equals(before, after)) changes[field] = new { from = before, to = after };
        }

        Value("Name", Name, edited.Name);
        Value("Description", Description, edited.Description);
        Value("Owner", Owner, edited.Owner);
        Value("Contact", Contact, edited.Contact);
        Value("IsOpenForTermCreation", IsOpenForTermCreation, edited.IsOpenForTermCreation);
        Value("IsAvailableForTagging", IsAvailableForTagging, edited.IsAvailableForTagging);

        var added = edited.Synonyms.Except(Synonyms, StringComparer.Ordinal).ToList();
        var removed = Synonyms.Except(edited.Synonyms, StringComparer.Ordinal).ToList();
        if (added.Count > 0 || removed.Count > 0) changes["Synonyms"] = new { added, removed, now = edited.Synonyms };

        PropertyChanges(changes, "CustomProperties", CustomProperties, edited.CustomProperties);
        PropertyChanges(changes, "LocalCustomProperties", LocalCustomProperties, edited.LocalCustomProperties);
        return changes;
    }

    private static void PropertyChanges(
        Dictionary<string, object> changes, string field, IReadOnlyList<PropertyEntry> before, IReadOnlyList<PropertyEntry> after)
    {
        var old = before.GroupBy(p => p.Key, StringComparer.Ordinal).ToDictionary(g => g.Key, g => g.Last().Value, StringComparer.Ordinal);
        var now = after.GroupBy(p => p.Key, StringComparer.Ordinal).ToDictionary(g => g.Key, g => g.Last().Value, StringComparer.Ordinal);
        var added = now.Where(p => !old.ContainsKey(p.Key)).ToDictionary(p => p.Key, p => p.Value);
        var removed = old.Where(p => !now.ContainsKey(p.Key)).ToDictionary(p => p.Key, p => p.Value);
        var changed = now.Where(p => old.TryGetValue(p.Key, out var v) && v != p.Value)
            .ToDictionary(p => p.Key, p => (object)new { from = old[p.Key], to = p.Value });
        if (added.Count > 0 || removed.Count > 0 || changed.Count > 0) changes[field] = new { added, removed, changed };
    }

    private static bool SameSet(IReadOnlyList<string> a, IReadOnlyList<string> b) =>
        a.Count == b.Count && a.ToHashSet(StringComparer.Ordinal).SetEquals(b);

    private static bool SameProperties(IReadOnlyList<PropertyEntry> a, IReadOnlyList<PropertyEntry> b) =>
        a.Count == b.Count && a.ToHashSet().SetEquals(b);
}
