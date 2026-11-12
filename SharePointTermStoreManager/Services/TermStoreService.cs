using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using SharePointTermStoreManager.Models;

namespace SharePointTermStoreManager.Services;

/// <summary>The default term store of a profile's site, with its term groups.</summary>
public sealed record TermStoreInfo(Guid Id, string Name, int DefaultLanguage, IReadOnlyList<TermNode> Groups);

/// <summary>Thrown when an item changed in SharePoint after it was opened, so saving would overwrite those changes.</summary>
public sealed class ConcurrentChangeException(string message) : Exception(message);

/// <summary>
/// Reads and changes the SharePoint Online term store through CSOM (<c>Microsoft.SharePoint.Client.Taxonomy</c>),
/// the only supported API that covers CRUD, merge, reuse and deprecate without an app registration. Requests
/// carry the same SharePoint access token <see cref="AuthService"/> gets for the profile's host.
///
/// A <see cref="ClientContext"/> isn't thread-safe, so every operation creates its own and disposes it. Changes
/// are sent in one batch ending in <see cref="TermStore.CommitAll"/>, so an operation either applies or fails as
/// a whole.
///
/// Terms are always addressed through their term set (<see cref="TermStore.GetTermInTermSet"/>): a reused term
/// has one ID but an instance in each term set, and changes must go to the instance shown in the tree.
/// </summary>
public sealed class TermStoreService(AuthService auth)
{
    /// <summary>Default language of each profile's term store, by profile ID.</summary>
    private readonly Dictionary<string, int> _languages = new();

    // ---- Reading -------------------------------------------------------------------------------------------

    /// <summary>Loads the default term store for the profile's site and its term groups.</summary>
    public async Task<TermStoreInfo> GetTermStoreAsync(ConnectionProfile profile, CancellationToken ct = default)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        ctx.Load(store,
            s => s.Id,
            s => s.Name,
            s => s.DefaultLanguage,
            s => s.Groups.Include(
                g => g.Id,
                g => g.Name,
                g => g.Description,
                g => g.IsSystemGroup,
                g => g.IsSiteCollectionGroup));
        await ctx.ExecuteQueryAsync();
        _languages[profile.Id] = store.DefaultLanguage;

        var groups = store.Groups.AsEnumerable()
            .Select(g => new TermNode
            {
                Kind = TermNodeKind.Group,
                Id = g.Id,
                Name = g.Name,
                Description = g.Description ?? "",
                IsSystemGroup = g.IsSystemGroup,
                IsSiteCollectionGroup = g.IsSiteCollectionGroup,
                HasUnrealizedChildren = true,
            })
            // System groups (e.g. "People", "Search Dictionaries") last, as the term store admin page shows them.
            .OrderBy(g => g.IsSystemGroup)
            .ThenBy(g => g.Name, StringComparer.CurrentCultureIgnoreCase)
            .ToList();

        return new TermStoreInfo(store.Id, store.Name, store.DefaultLanguage, groups);
    }

    /// <summary>Loads the direct children of a node: a group's term sets, or a term set's / term's terms.</summary>
    public Task<IReadOnlyList<TermNode>> GetChildrenAsync(ConnectionProfile profile, TermNode parent, CancellationToken ct = default) =>
        parent.Kind == TermNodeKind.Group
            ? GetTermSetsAsync(profile, parent, ct)
            : GetTermsAsync(profile, parent, ct);

    private async Task<IReadOnlyList<TermNode>> GetTermSetsAsync(ConnectionProfile profile, TermNode group, CancellationToken ct)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var termSets = GetStore(ctx).GetGroup(group.Id).TermSets;
        ctx.Load(termSets, c => c.Include(t => t.Id, t => t.Name, t => t.Description));
        await ctx.ExecuteQueryAsync();

        return termSets.AsEnumerable()
            .Select(t => new TermNode
            {
                Kind = TermNodeKind.TermSet,
                Id = t.Id,
                Name = t.Name,
                Description = t.Description ?? "",
                Parent = group,
                // Term sets don't report a term count, so assume children; expanding an empty one clears the chevron.
                HasUnrealizedChildren = true,
            })
            .OrderBy(t => t.Name, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    private async Task<IReadOnlyList<TermNode>> GetTermsAsync(ConnectionProfile profile, TermNode parent, CancellationToken ct)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);

        TermSetItem item = parent.Kind == TermNodeKind.TermSet ? store.GetTermSet(parent.Id) : GetTerm(store, parent);
        ctx.Load(item, t => t.CustomSortOrder);
        ctx.Load(item.Terms, TermFields);
        await ctx.ExecuteQueryAsync();

        parent.HasCustomSortOrder = !string.IsNullOrWhiteSpace(item.CustomSortOrder);
        return ToNodes(item.Terms, parent, item.CustomSortOrder);
    }

    private static readonly System.Linq.Expressions.Expression<Func<TermCollection, object>>[] TermFields =
    [
        c => c.Include(
            t => t.Id,
            t => t.Name,
            t => t.Description,
            t => t.IsDeprecated,
            t => t.IsReused,
            t => t.IsPinned,
            t => t.IsSourceTerm,
            t => t.TermsCount),
    ];

    /// <summary>
    /// Term nodes in the order the term store admin page shows them: the parent's custom sort order when it has
    /// one (a colon-separated list of term IDs), otherwise alphabetical.
    /// </summary>
    private static IReadOnlyList<TermNode> ToNodes(TermCollection terms, TermNode parent, string? customSortOrder)
    {
        var order = (customSortOrder ?? "")
            .Split(':', StringSplitOptions.RemoveEmptyEntries)
            .Select((id, index) => (ok: Guid.TryParse(id, out var guid), guid, index))
            .Where(x => x.ok)
            .GroupBy(x => x.guid)
            .ToDictionary(g => g.Key, g => g.First().index);

        return terms.AsEnumerable()
            .Select(t => new TermNode
            {
                Kind = TermNodeKind.Term,
                Id = t.Id,
                Name = t.Name,
                Description = t.Description ?? "",
                Parent = parent,
                IsDeprecated = t.IsDeprecated,
                IsReused = t.IsReused,
                IsPinned = t.IsPinned,
                IsSourceTerm = t.IsSourceTerm,
                ChildTermCount = t.TermsCount,
                HasUnrealizedChildren = t.TermsCount > 0,
            })
            .OrderBy(n => order.TryGetValue(n.Id, out var i) ? i : int.MaxValue)
            .ThenBy(n => n.Name, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    /// <summary>Reads the editable properties of a node, in the term store's default language.</summary>
    public async Task<NodeDetails> GetDetailsAsync(ConnectionProfile profile, TermNode node, CancellationToken ct = default)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        int lcid = await GetLanguageAsync(profile, ctx, store);

        switch (node.Kind)
        {
            case TermNodeKind.Group:
            {
                var group = store.GetGroup(node.Id);
                ctx.Load(group,
                    g => g.Id, g => g.Name, g => g.Description, g => g.IsSystemGroup,
                    g => g.CreatedDate, g => g.LastModifiedDate);
                await ctx.ExecuteQueryAsync();
                return new NodeDetails
                {
                    Kind = node.Kind,
                    Id = group.Id,
                    Language = lcid,
                    Name = group.Name,
                    Description = group.Description ?? "",
                    IsSystemGroup = group.IsSystemGroup,
                    Created = group.CreatedDate,
                    LastModified = group.LastModifiedDate,
                };
            }

            case TermNodeKind.TermSet:
            {
                var termSet = store.GetTermSet(node.Id);
                ctx.Load(termSet,
                    t => t.Id, t => t.Name, t => t.Description, t => t.Owner, t => t.Contact,
                    t => t.IsAvailableForTagging, t => t.IsOpenForTermCreation, t => t.CustomProperties,
                    t => t.CustomSortOrder, t => t.CreatedDate, t => t.LastModifiedDate);
                await ctx.ExecuteQueryAsync();
                return new NodeDetails
                {
                    Kind = node.Kind,
                    Id = termSet.Id,
                    Language = lcid,
                    Name = termSet.Name,
                    Description = termSet.Description ?? "",
                    Owner = termSet.Owner ?? "",
                    Contact = termSet.Contact ?? "",
                    IsAvailableForTagging = termSet.IsAvailableForTagging,
                    IsOpenForTermCreation = termSet.IsOpenForTermCreation,
                    HasCustomSortOrder = !string.IsNullOrWhiteSpace(termSet.CustomSortOrder),
                    CustomProperties = ToEntries(termSet.CustomProperties),
                    Created = termSet.CreatedDate,
                    LastModified = termSet.LastModifiedDate,
                };
            }

            default:
            {
                var term = GetTerm(store, node);
                ctx.Load(term,
                    t => t.Id, t => t.Name, t => t.IsAvailableForTagging,
                    t => t.CustomProperties, t => t.LocalCustomProperties,
                    t => t.IsDeprecated, t => t.IsReused, t => t.IsPinned, t => t.IsSourceTerm, t => t.TermsCount,
                    t => t.CreatedDate, t => t.LastModifiedDate, t => t.CustomSortOrder,
                    t => t.Labels.Include(l => l.Value, l => l.Language, l => l.IsDefaultForLanguage),
                    t => t.TermSets.Include(s => s.Id, s => s.Name, s => s.Description));
                var description = term.GetDescription(lcid);
                await ctx.ExecuteQueryAsync();

                var labels = term.Labels.AsEnumerable().Where(l => l.Language == lcid).ToList();
                return new NodeDetails
                {
                    Kind = node.Kind,
                    Id = term.Id,
                    Language = lcid,
                    Name = labels.FirstOrDefault(l => l.IsDefaultForLanguage)?.Value ?? term.Name,
                    Description = description.Value ?? "",
                    Synonyms = labels.Where(l => !l.IsDefaultForLanguage).Select(l => l.Value).ToList(),
                    OtherLanguages = OtherLanguageLabels(term.Labels, lcid),
                    MemberOf = await GetMembershipAsync(ctx, store, term, node),
                    HasCustomSortOrder = !string.IsNullOrWhiteSpace(term.CustomSortOrder),
                    IsAvailableForTagging = term.IsAvailableForTagging,
                    CustomProperties = ToEntries(term.CustomProperties),
                    LocalCustomProperties = ToEntries(term.LocalCustomProperties),
                    IsDeprecated = term.IsDeprecated,
                    IsReused = term.IsReused,
                    IsPinned = term.IsPinned,
                    IsSourceTerm = term.IsSourceTerm,
                    TermsCount = term.TermsCount,
                    Created = term.CreatedDate,
                    LastModified = term.LastModifiedDate,
                };
            }
        }
    }

    /// <summary>
    /// The term sets a term is in, as the term store's "Member of" list: the one it's selected in first, then the
    /// others alphabetically, each with whether that copy is the source, a reused copy or a pinned copy. A term
    /// that isn't reused is in one term set, as its source; otherwise each copy is read in one more request.
    /// </summary>
    private static async Task<IReadOnlyList<TermSetMembership>> GetMembershipAsync(ClientContext ctx, TermStore store, Term term, TermNode node)
    {
        var termSets = term.TermSets.AsEnumerable()
            .OrderBy(s => s.Id == node.TermSet?.Id ? 0 : 1)
            .ThenBy(s => s.Name, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
        if (!term.IsReused)
            return termSets.Select(s => new TermSetMembership(s.Name, s.Description ?? "", "Source")).ToList();

        var copies = termSets.Select(s =>
        {
            var copy = store.GetTermInTermSet(s.Id, term.Id);
            ctx.Load(copy, t => t.IsSourceTerm, t => t.IsPinned);
            return (termSet: s, copy);
        }).ToList();
        await ctx.ExecuteQueryAsync();

        return copies
            .Select(c => new TermSetMembership(
                c.termSet.Name,
                c.termSet.Description ?? "",
                c.copy.IsSourceTerm ? "Source" : c.copy.IsPinned ? "Pinned" : "Reused"))
            .ToList();
    }

    /// <summary>The default label and synonyms in each language other than <paramref name="lcid"/>.</summary>
    private static List<LanguageLabels> OtherLanguageLabels(LabelCollection labels, int lcid) =>
        labels.AsEnumerable()
            .Where(l => l.Language != lcid)
            .GroupBy(l => l.Language)
            .Select(g => new LanguageLabels(
                g.Key,
                LanguageName(g.Key),
                g.FirstOrDefault(l => l.IsDefaultForLanguage)?.Value ?? "",
                string.Join("; ", g.Where(l => !l.IsDefaultForLanguage).Select(l => l.Value))))
            .OrderBy(l => l.Language, StringComparer.CurrentCultureIgnoreCase)
            .ToList();

    /// <summary>"English (United States)" for 1033.</summary>
    public static string LanguageName(int lcid)
    {
        try { return System.Globalization.CultureInfo.GetCultureInfo(lcid).DisplayName; }
        catch (System.Globalization.CultureNotFoundException) { return $"LCID {lcid}"; }
    }

    private static List<PropertyEntry> ToEntries(IDictionary<string, string>? properties) =>
        (properties ?? new Dictionary<string, string>())
        .Select(p => new PropertyEntry(p.Key, p.Value ?? ""))
        .OrderBy(p => p.Key, StringComparer.CurrentCultureIgnoreCase)
        .ToList();

    // ---- Exporting -----------------------------------------------------------------------------------------

    /// <summary>
    /// Reads the term store for an export: the term groups (all, or those in <paramref name="groupIds"/>), their
    /// term sets and every term with its details, ordered as the tree shows them. One request reads the groups and
    /// term sets, then one per term set reads all its terms (<see cref="TermSet.GetAllTerms"/>); the hierarchy is
    /// rebuilt from each term's path. With <paramref name="includeDeprecated"/> false, deprecated terms and the
    /// terms below them are left out.
    /// </summary>
    public async Task<ExportStore> ReadForExportAsync(
        ConnectionProfile profile, IReadOnlySet<Guid>? groupIds, bool includeDeprecated,
        IProgress<ExportProgress>? progress, CancellationToken ct = default)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        progress?.Report(new ExportProgress(0, 0, "Reading term groups and term sets…"));
        ctx.Load(store,
            s => s.Id, s => s.Name, s => s.DefaultLanguage,
            s => s.Groups.Include(
                g => g.Id, g => g.Name, g => g.Description, g => g.IsSystemGroup, g => g.IsSiteCollectionGroup,
                g => g.CreatedDate, g => g.LastModifiedDate,
                g => g.TermSets.Include(
                    t => t.Id, t => t.Name, t => t.Description, t => t.Owner, t => t.Contact,
                    t => t.IsOpenForTermCreation, t => t.IsAvailableForTagging, t => t.CustomProperties,
                    t => t.CustomSortOrder, t => t.CreatedDate, t => t.LastModifiedDate)));
        await ctx.ExecuteQueryAsync();
        ct.ThrowIfCancellationRequested();
        int lcid = store.DefaultLanguage;

        var groups = store.Groups.AsEnumerable()
            .Where(g => groupIds == null || groupIds.Contains(g.Id))
            .OrderBy(g => g.IsSystemGroup)
            .ThenBy(g => g.Name, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
        int total = groups.Sum(g => g.TermSets.Count), done = 0;

        var result = new ExportStore
        {
            Id = store.Id,
            Name = store.Name,
            Site = profile.SiteUrl,
            Exported = DateTimeOffset.Now,
            DefaultLanguage = lcid,
        };
        foreach (var group in groups)
        {
            var exportGroup = new ExportGroup
            {
                Id = group.Id,
                Name = group.Name,
                Description = group.Description ?? "",
                IsSystemGroup = group.IsSystemGroup,
                IsSiteCollectionGroup = group.IsSiteCollectionGroup,
                Created = group.CreatedDate,
                LastModified = group.LastModifiedDate,
            };
            result.Groups.Add(exportGroup);

            foreach (var termSet in group.TermSets.AsEnumerable().OrderBy(t => t.Name, StringComparer.CurrentCultureIgnoreCase))
            {
                ct.ThrowIfCancellationRequested();
                progress?.Report(new ExportProgress(done, total, $"Reading {group.Name} › {termSet.Name}…"));

                var terms = termSet.GetAllTerms();
                ctx.Load(terms, c => c.Include(
                    t => t.Id, t => t.Name, t => t.Description, t => t.PathOfTerm,
                    t => t.IsDeprecated, t => t.IsAvailableForTagging, t => t.IsReused, t => t.IsPinned, t => t.IsSourceTerm,
                    t => t.CustomSortOrder, t => t.CustomProperties, t => t.LocalCustomProperties,
                    t => t.CreatedDate, t => t.LastModifiedDate,
                    t => t.Labels.Include(l => l.Value, l => l.Language, l => l.IsDefaultForLanguage)));
                await ctx.ExecuteQueryAsync();
                ct.ThrowIfCancellationRequested();

                exportGroup.TermSets.Add(new ExportTermSet
                {
                    Id = termSet.Id,
                    Name = termSet.Name,
                    Description = termSet.Description ?? "",
                    Owner = termSet.Owner ?? "",
                    Contact = termSet.Contact ?? "",
                    IsOpenForTermCreation = termSet.IsOpenForTermCreation,
                    IsAvailableForTagging = termSet.IsAvailableForTagging,
                    HasCustomSortOrder = !string.IsNullOrWhiteSpace(termSet.CustomSortOrder),
                    CustomProperties = ToDictionary(termSet.CustomProperties),
                    Created = termSet.CreatedDate,
                    LastModified = termSet.LastModifiedDate,
                    Terms = BuildTermTree(terms, termSet.CustomSortOrder, lcid, includeDeprecated),
                });
                done++;
            }
        }
        progress?.Report(new ExportProgress(done, total, "Read the term store."));
        return result;
    }

    /// <summary>
    /// Nests a term set's terms (all levels, as <see cref="TermSet.GetAllTerms"/> returns them) by their paths:
    /// <see cref="Term.PathOfTerm"/> is the default labels from the top down, ";"-separated, and sibling labels are
    /// unique, so it identifies the parent. Each level is sorted as the tree sorts it.
    /// </summary>
    private static List<ExportTerm> BuildTermTree(TermCollection terms, string? termSetSortOrder, int lcid, bool includeDeprecated)
    {
        var nodes = terms.AsEnumerable().Select(t => (path: t.PathOfTerm ?? t.Name, node: ToExportTerm(t, lcid))).ToList();
        var byPath = new Dictionary<string, ExportTerm>(StringComparer.Ordinal);
        foreach (var (path, node) in nodes) byPath.TryAdd(path, node);

        var roots = new List<ExportTerm>();
        foreach (var (path, node) in nodes)
        {
            int cut = path.LastIndexOf(';');
            if (cut > 0 && byPath.TryGetValue(path[..cut], out var parent) && parent != node) parent.Terms.Add(node);
            else roots.Add(node);
        }

        List<ExportTerm> Arrange(List<ExportTerm> level, string? sortOrder)
        {
            var order = SortIndex(sortOrder);
            var arranged = level
                .Where(t => includeDeprecated || !t.IsDeprecated)
                .OrderBy(t => order.TryGetValue(t.Id, out var i) ? i : int.MaxValue)
                .ThenBy(t => t.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
            foreach (var term in arranged) term.Terms = Arrange(term.Terms, term.CustomSortOrder);
            return arranged;
        }
        return Arrange(roots, termSetSortOrder);
    }

    private static ExportTerm ToExportTerm(Term t, int lcid)
    {
        var labels = t.Labels.AsEnumerable().ToList();
        var own = labels.Where(l => l.Language == lcid).ToList();
        var others = OtherLanguageLabels(t.Labels, lcid)
            .Select(l => new ExportLanguageLabels(
                l.Lcid, l.Language, l.Label,
                l.Synonyms.Split("; ", StringSplitOptions.RemoveEmptyEntries).ToList()))
            .ToList();
        return new ExportTerm
        {
            Id = t.Id,
            Name = t.Name,
            Description = t.Description ?? "",
            Synonyms = own.Where(l => !l.IsDefaultForLanguage).Select(l => l.Value).ToList(),
            OtherLanguages = others.Count > 0 ? others : null,
            IsDeprecated = t.IsDeprecated,
            IsAvailableForTagging = t.IsAvailableForTagging,
            IsReused = t.IsReused,
            IsPinned = t.IsPinned,
            IsSourceTerm = t.IsSourceTerm,
            HasCustomSortOrder = !string.IsNullOrWhiteSpace(t.CustomSortOrder),
            CustomSortOrder = t.CustomSortOrder,
            CustomProperties = ToDictionary(t.CustomProperties),
            LocalCustomProperties = ToDictionary(t.LocalCustomProperties),
            Created = t.CreatedDate,
            LastModified = t.LastModifiedDate,
        };
    }

    private static Dictionary<string, string> ToDictionary(IDictionary<string, string>? properties) =>
        ToEntries(properties).ToDictionary(p => p.Key, p => p.Value);

    /// <summary>Positions of term IDs in a custom sort order (a colon-separated list of IDs).</summary>
    private static Dictionary<Guid, int> SortIndex(string? customSortOrder) =>
        (customSortOrder ?? "")
        .Split(':', StringSplitOptions.RemoveEmptyEntries)
        .Select((id, index) => (ok: Guid.TryParse(id, out var guid), guid, index))
        .Where(x => x.ok)
        .GroupBy(x => x.guid)
        .ToDictionary(g => g.Key, g => g.First().index);

    // ---- Changing ------------------------------------------------------------------------------------------

    /// <summary>
    /// Creates a term group (<paramref name="parent"/> null), a term set in a group, or a term in a term set or
    /// under a term. <paramref name="name"/> must already be normalized (<see cref="Helpers.TermNames"/>).
    /// Returns the new node, not yet added to the tree.
    /// </summary>
    public async Task<TermNode> CreateAsync(
        ConnectionProfile profile, TermNode? parent, TermNodeKind kind, string name, string description, CancellationToken ct = default)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        int lcid = await GetLanguageAsync(profile, ctx, store);
        var id = Guid.NewGuid();

        TaxonomyItem created;
        switch (kind)
        {
            case TermNodeKind.Group:
            {
                var group = store.CreateGroup(name, id);
                if (description.Length > 0) group.Description = description;
                created = group;
                break;
            }
            case TermNodeKind.TermSet:
            {
                if (parent is not { Kind: TermNodeKind.Group }) throw new ArgumentException("A term set must be created in a term group.");
                var termSet = store.GetGroup(parent.Id).CreateTermSet(name, id, lcid);
                if (description.Length > 0) termSet.Description = description;
                created = termSet;
                break;
            }
            default:
            {
                if (parent is null or { Kind: TermNodeKind.Group }) throw new ArgumentException("A term must be created in a term set or under a term.");
                TermSetItem container = parent.Kind == TermNodeKind.TermSet ? store.GetTermSet(parent.Id) : GetTerm(store, parent);
                var term = container.CreateTerm(name, lcid, id);
                if (description.Length > 0) term.SetDescription(description, lcid);
                created = term;
                break;
            }
        }

        store.CommitAll();
        ctx.Load(created, c => c.Name); // as SharePoint stored it
        await ctx.ExecuteQueryAsync();

        return new TermNode
        {
            Kind = kind,
            Id = id,
            Name = created.Name,
            Description = description,
            Parent = parent,
            HasUnrealizedChildren = false, // brand new, so known to be empty
        };
    }

    /// <summary>
    /// Saves the differences between <paramref name="original"/> (as read by <see cref="GetDetailsAsync"/>) and
    /// <paramref name="edited"/>. Fails with <see cref="ConcurrentChangeException"/> when the item changed in
    /// SharePoint since <paramref name="original"/> was read.
    /// </summary>
    public async Task UpdateAsync(
        ConnectionProfile profile, TermNode node, NodeDetails original, NodeDetails edited, CancellationToken ct = default)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        int lcid = original.Language;

        switch (node.Kind)
        {
            case TermNodeKind.Group:
            {
                var group = store.GetGroup(node.Id);
                ctx.Load(group, g => g.LastModifiedDate);
                await ctx.ExecuteQueryAsync();
                EnsureUnchanged(node, original, group.LastModifiedDate);

                if (edited.Name != original.Name) group.Name = edited.Name;
                if (edited.Description != original.Description) group.Description = edited.Description;
                break;
            }

            case TermNodeKind.TermSet:
            {
                var termSet = store.GetTermSet(node.Id);
                ctx.Load(termSet, t => t.LastModifiedDate);
                await ctx.ExecuteQueryAsync();
                EnsureUnchanged(node, original, termSet.LastModifiedDate);

                if (edited.Name != original.Name) termSet.Name = edited.Name;
                if (edited.Description != original.Description) termSet.Description = edited.Description;
                if (edited.Owner != original.Owner) termSet.Owner = edited.Owner;
                if (edited.Contact != original.Contact) termSet.Contact = edited.Contact;
                if (edited.IsAvailableForTagging != original.IsAvailableForTagging) termSet.IsAvailableForTagging = edited.IsAvailableForTagging;
                if (edited.IsOpenForTermCreation != original.IsOpenForTermCreation) termSet.IsOpenForTermCreation = edited.IsOpenForTermCreation;
                ApplyProperties(original.CustomProperties, edited.CustomProperties, termSet.DeleteCustomProperty, termSet.SetCustomProperty);
                break;
            }

            default:
            {
                var term = GetTerm(store, node);
                ctx.Load(term, t => t.LastModifiedDate);
                ctx.Load(term.Labels, c => c.Include(l => l.Value, l => l.Language, l => l.IsDefaultForLanguage));
                await ctx.ExecuteQueryAsync();
                EnsureUnchanged(node, original, term.LastModifiedDate);

                ApplyLabels(term, lcid, edited.Name, edited.Synonyms);
                if (edited.Description != original.Description) term.SetDescription(edited.Description, lcid);
                if (edited.IsAvailableForTagging != original.IsAvailableForTagging) term.IsAvailableForTagging = edited.IsAvailableForTagging;
                ApplyProperties(original.CustomProperties, edited.CustomProperties, term.DeleteCustomProperty, term.SetCustomProperty);
                ApplyProperties(original.LocalCustomProperties, edited.LocalCustomProperties, term.DeleteLocalCustomProperty, term.SetLocalCustomProperty);
                break;
            }
        }

        store.CommitAll();
        await ctx.ExecuteQueryAsync();
    }

    /// <summary>
    /// Deletes a term group (which must be empty), a term set with all its terms, or a term with all the terms
    /// below it. A term is deleted only from the term set shown in the tree.
    /// </summary>
    public async Task DeleteAsync(ConnectionProfile profile, TermNode node, CancellationToken ct = default)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);

        switch (node.Kind)
        {
            case TermNodeKind.Group:
            {
                var group = store.GetGroup(node.Id);
                ctx.Load(group.TermSets, c => c.Include(t => t.Id));
                await ctx.ExecuteQueryAsync();
                if (group.TermSets.Count > 0)
                    throw new InvalidOperationException(
                        $"“{node.Name}” still contains {group.TermSets.Count} term set(s). Delete or move them first; SharePoint only deletes empty term groups.");
                group.DeleteObject();
                break;
            }
            case TermNodeKind.TermSet:
                store.GetTermSet(node.Id).DeleteObject();
                break;
            default:
                GetTerm(store, node).DeleteObject();
                break;
        }

        store.CommitAll();
        await ctx.ExecuteQueryAsync();
    }

    private static void EnsureUnchanged(TermNode node, NodeDetails original, DateTime lastModified)
    {
        if (lastModified != original.LastModified)
            throw new ConcurrentChangeException(
                $"“{node.Name}” was changed in SharePoint after you opened it. Nothing was saved. Discard your changes (or copy them), then select the item again to load the current version.");
    }

    /// <summary>
    /// Updates a term's labels in one language: the default label becomes <paramref name="name"/> and the other
    /// labels become <paramref name="synonyms"/>. Labels in other languages aren't touched. Renaming to an existing
    /// synonym promotes that label, so the old name becomes a synonym unless it's also removed.
    /// </summary>
    private static void ApplyLabels(Term term, int lcid, string name, IReadOnlyList<string> synonyms)
    {
        var labels = term.Labels.AsEnumerable().Where(l => l.Language == lcid).ToList();
        var defaultLabel = labels.FirstOrDefault(l => l.IsDefaultForLanguage);
        var others = labels.Where(l => !l.IsDefaultForLanguage).ToList();

        if (defaultLabel == null)
        {
            term.Name = name;
        }
        else if (defaultLabel.Value != name)
        {
            var promoted = others.FirstOrDefault(l => string.Equals(l.Value, name, StringComparison.OrdinalIgnoreCase));
            if (promoted != null)
            {
                promoted.SetAsDefaultForLanguage();
                if (promoted.Value != name) promoted.Value = name;
                others.Remove(promoted);
                others.Add(defaultLabel); // the old default label is now an ordinary label
            }
            else
            {
                term.Name = name;
            }
        }

        // Remove labels that are no longer wanted first, then change case or add, so a label that only changes
        // case is never briefly duplicated.
        var wanted = synonyms.ToList();
        var kept = new List<(Label label, string value)>();
        foreach (var label in others)
        {
            var match = wanted.FirstOrDefault(w => string.Equals(w, label.Value, StringComparison.OrdinalIgnoreCase));
            if (match == null)
            {
                label.DeleteObject();
                continue;
            }
            wanted.Remove(match);
            kept.Add((label, match));
        }
        foreach (var (label, value) in kept)
        {
            if (label.Value != value) label.Value = value;
        }
        foreach (var synonym in wanted) term.CreateLabel(synonym, lcid, false);
    }

    private static void ApplyProperties(
        IReadOnlyList<PropertyEntry> before, IReadOnlyList<PropertyEntry> after, Action<string> delete, Action<string, string> set)
    {
        var old = before.ToDictionary(p => p.Key, p => p.Value, StringComparer.Ordinal);
        var keys = after.Select(p => p.Key).ToHashSet(StringComparer.Ordinal);

        foreach (var key in old.Keys.Where(k => !keys.Contains(k))) delete(key);
        foreach (var (key, value) in after)
        {
            if (!old.TryGetValue(key, out var oldValue) || oldValue != value) set(key, value);
        }
    }

    // ---- Searching -----------------------------------------------------------------------------------------

    /// <summary>
    /// Terms with a label (default label or synonym) starting with <paramref name="text"/>, in one term set or the
    /// whole term store. CSOM label search only supports "starts with" and exact matches.
    /// </summary>
    public async Task<IReadOnlyList<ItemRef>> SearchTermsAsync(
        ConnectionProfile profile, string text, Guid? termSetId, int maxResults = 100, CancellationToken ct = default)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        int lcid = await GetLanguageAsync(profile, ctx, store);

        var match = new LabelMatchInformation(ctx)
        {
            TermLabel = text,
            StringMatchOption = StringMatchOption.StartsWith,
            ResultCollectionSize = maxResults,
            Lcid = lcid,
            DefaultLabelOnly = false,
            TrimDeprecated = false,
            TrimUnavailable = false,
            ExcludeKeyword = false,
        };
        var terms = termSetId is { } id ? store.GetTermSet(id).GetTerms(match) : store.GetTerms(match);
        ctx.Load(terms, c => c.Include(TermRefFields));
        await ctx.ExecuteQueryAsync();
        ct.ThrowIfCancellationRequested();

        return terms.AsEnumerable()
            .Select(TermRef)
            .OrderBy(r => r.Name, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(r => r.Path, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    /// <summary>
    /// The term group, term set or term with this ID (unique identifier). A reused term has the same ID in every
    /// term set it's in, so it can give several results, one per term set.
    /// </summary>
    public async Task<IReadOnlyList<ItemRef>> FindByIdAsync(ConnectionProfile profile, Guid id, CancellationToken ct = default)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        var results = new List<ItemRef>();

        // One request per kind: an ID that isn't a group (say) can fail its request without hiding the others.
        var group = store.GetGroup(id);
        ctx.Load(group, g => g.Id, g => g.Name, g => g.IsSystemGroup);
        if (await TryExecuteAsync(ctx, group, ct))
            results.Add(new ItemRef(TermNodeKind.Group, group.Id, null, group.Name, group.Name) { IsSystemGroup = group.IsSystemGroup });

        var termSet = store.GetTermSet(id);
        ctx.Load(termSet, t => t.Id, t => t.Name, t => t.Group.Id, t => t.Group.Name);
        if (await TryExecuteAsync(ctx, termSet, ct))
            results.Add(new ItemRef(TermNodeKind.TermSet, termSet.Id, termSet.Id, termSet.Name, $"{termSet.Group.Name} › {termSet.Name}")
            {
                GroupId = termSet.Group.Id,
            });

        var term = store.GetTerm(id);
        ctx.Load(term, TermRefFields);
        ctx.Load(term.ReusedTerms, c => c.Include(TermRefFields));
        if (await TryExecuteAsync(ctx, term, ct))
        {
            results.AddRange(new[] { term }.Concat(term.ReusedTerms)
                .Select(TermRef)
                .GroupBy(r => r.TermSetId).Select(g => g.First())
                .OrderBy(r => r.Path, StringComparer.CurrentCultureIgnoreCase));
        }
        return results;
    }

    /// <summary>Runs the pending request; false when the object it loads doesn't exist.</summary>
    private static async Task<bool> TryExecuteAsync(ClientContext ctx, ClientObject loaded, CancellationToken ct)
    {
        try
        {
            await ctx.ExecuteQueryAsync();
        }
        catch (ServerException)
        {
            return false;
        }
        ct.ThrowIfCancellationRequested();
        return loaded.ServerObjectIsNull != true;
    }

    /// <summary>The term properties <see cref="TermRef"/> reads.</summary>
    private static readonly System.Linq.Expressions.Expression<Func<Term, object>>[] TermRefFields =
    [
        t => t.Id, t => t.Name, t => t.PathOfTerm,
        t => t.IsDeprecated, t => t.IsReused, t => t.IsPinned, t => t.IsSourceTerm, t => t.TermsCount,
        t => t.TermSet.Id, t => t.TermSet.Name, t => t.TermSet.Group.Id, t => t.TermSet.Group.Name,
    ];

    /// <summary>A search result for a term loaded with <see cref="TermRefFields"/>.</summary>
    private static ItemRef TermRef(Term t)
    {
        // PathOfTerm is the default labels from the top of the term set down to the term, ";"-separated.
        var labels = (t.PathOfTerm ?? t.Name).Split(';', StringSplitOptions.RemoveEmptyEntries);
        return new ItemRef(
            TermNodeKind.Term, t.Id, t.TermSet.Id, t.Name,
            string.Join(" › ", new[] { t.TermSet.Group.Name, t.TermSet.Name }.Concat(labels)))
        {
            IsDeprecated = t.IsDeprecated,
            IsReused = t.IsReused,
            IsPinned = t.IsPinned,
            IsSourceTerm = t.IsSourceTerm,
            ChildTermCount = t.TermsCount,
            GroupId = t.TermSet.Group.Id,
            ParentLabels = labels.Length > 1 ? labels[..^1] : [],
        };
    }

    /// <summary>Term sets named exactly <paramref name="name"/> (CSOM can't search term set names by prefix).</summary>
    public async Task<IReadOnlyList<ItemRef>> FindTermSetsByNameAsync(ConnectionProfile profile, string name, CancellationToken ct = default)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        int lcid = await GetLanguageAsync(profile, ctx, store);
        var termSets = store.GetTermSetsByName(name, lcid);
        ctx.Load(termSets, c => c.Include(t => t.Id, t => t.Name, t => t.Group.Id, t => t.Group.Name));
        await ctx.ExecuteQueryAsync();
        ct.ThrowIfCancellationRequested();

        return termSets.AsEnumerable()
            .Select(t => new ItemRef(TermNodeKind.TermSet, t.Id, t.Id, t.Name, $"{t.Group.Name} › {t.Name}") { GroupId = t.Group.Id })
            .ToList();
    }

    // ---- Term operations -----------------------------------------------------------------------------------

    /// <summary>Deprecates or un-deprecates a term. Returns whether it's now deprecated.</summary>
    public async Task<bool> SetDeprecatedAsync(ConnectionProfile profile, ItemRef term, bool deprecate, CancellationToken ct = default)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        var t = GetTerm(store, term);
        t.Deprecate(deprecate);
        store.CommitAll();
        ctx.Load(t, x => x.IsDeprecated);
        await ctx.ExecuteQueryAsync();
        return t.IsDeprecated;
    }

    /// <summary>
    /// Merges <paramref name="source"/> into <paramref name="target"/>: the source's labels become labels of the
    /// target, its ID is recorded on the target as a merged term ID, and the source is removed. Terms with child
    /// terms aren't merged; their children have to be moved or merged first.
    /// </summary>
    public async Task MergeAsync(ConnectionProfile profile, ItemRef source, ItemRef target, CancellationToken ct = default)
    {
        if (source.Id == target.Id) throw new InvalidOperationException("A term can't be merged into itself.");

        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        var from = GetTerm(store, source);
        var into = GetTerm(store, target);
        ctx.Load(from, t => t.TermsCount, t => t.Name);
        ctx.Load(into, t => t.IsPinned, t => t.IsSourceTerm, t => t.Name);
        var intoIsBelow = into.GetIsDescendantOf(from);
        await ctx.ExecuteQueryAsync();

        if (from.TermsCount > 0)
            throw new InvalidOperationException(
                $"“{from.Name}” has {from.TermsCount} child term(s). Move or merge them first, so it's clear where they end up.");
        if (intoIsBelow.Value)
            throw new InvalidOperationException($"“{into.Name}” is below “{from.Name}”, so it can't be the merge target.");
        if (into.IsPinned && !into.IsSourceTerm)
            throw new InvalidOperationException($"“{into.Name}” is a pinned copy. Pick the term in its source term set instead.");

        from.Merge(into);
        store.CommitAll();
        await ctx.ExecuteQueryAsync();
    }

    /// <summary>Moves a term (with everything below it) into a term set or under another term.</summary>
    public async Task MoveTermAsync(ConnectionProfile profile, ItemRef term, ItemRef newParent, CancellationToken ct = default)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        var t = GetTerm(store, term);

        TermSetItem destination;
        if (newParent.Kind == TermNodeKind.Term)
        {
            if (newParent.Id == term.Id) throw new InvalidOperationException("A term can't be moved under itself.");
            var parentTerm = GetTerm(store, newParent);
            var isBelow = parentTerm.GetIsDescendantOf(t);
            await ctx.ExecuteQueryAsync();
            if (isBelow.Value) throw new InvalidOperationException($"“{newParent.Name}” is below “{term.Name}”, so “{term.Name}” can't be moved there.");
            destination = parentTerm;
        }
        else if (newParent.Kind == TermNodeKind.TermSet)
        {
            destination = store.GetTermSet(newParent.Id);
        }
        else
        {
            throw new ArgumentException("A term can only be moved into a term set or under a term.");
        }

        t.Move(destination);
        store.CommitAll();
        await ctx.ExecuteQueryAsync();
    }

    /// <summary>Moves a term set, with all its terms, to another term group.</summary>
    public async Task MoveTermSetAsync(ConnectionProfile profile, ItemRef termSet, ItemRef group, CancellationToken ct = default)
    {
        if (group.Kind != TermNodeKind.Group) throw new ArgumentException("A term set can only be moved to a term group.");
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        store.GetTermSet(termSet.Id).Move(store.GetGroup(group.Id));
        store.CommitAll();
        await ctx.ExecuteQueryAsync();
    }

    /// <summary>
    /// Copies a term next to itself (SharePoint names the copy), optionally with its child terms. Returns the copy,
    /// as a node under <paramref name="parent"/>.
    /// </summary>
    public async Task<TermNode> CopyTermAsync(
        ConnectionProfile profile, ItemRef term, TermNode parent, bool includeChildren, CancellationToken ct = default)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        var copy = GetTerm(store, term).Copy(includeChildren);
        store.CommitAll();
        ctx.Load(copy, t => t.Id, t => t.Name, t => t.Description, t => t.IsDeprecated, t => t.IsReused,
            t => t.IsPinned, t => t.IsSourceTerm, t => t.TermsCount);
        await ctx.ExecuteQueryAsync();

        return new TermNode
        {
            Kind = TermNodeKind.Term,
            Id = copy.Id,
            Name = copy.Name,
            Description = copy.Description ?? "",
            Parent = parent,
            IsDeprecated = copy.IsDeprecated,
            IsReused = copy.IsReused,
            IsPinned = copy.IsPinned,
            IsSourceTerm = copy.IsSourceTerm,
            ChildTermCount = copy.TermsCount,
            HasUnrealizedChildren = copy.TermsCount > 0,
        };
    }

    /// <summary>
    /// Reuses <paramref name="source"/> in a term set or under a term. With <paramref name="pin"/>, the reused
    /// copy (and its children) is pinned: it can only be changed at the source. Otherwise
    /// <paramref name="includeChildren"/> reuses the whole branch.
    /// </summary>
    public async Task ReuseAsync(
        ConnectionProfile profile, ItemRef source, ItemRef destination, bool pin, bool includeChildren, CancellationToken ct = default)
    {
        using var ctx = await CreateContextAsync(profile, ct);
        var store = GetStore(ctx);
        var term = GetTerm(store, source);

        TermSetItem target = destination.Kind switch
        {
            TermNodeKind.TermSet => store.GetTermSet(destination.Id),
            TermNodeKind.Term => GetTerm(store, destination),
            _ => throw new ArgumentException("A term can only be reused in a term set or under a term."),
        };

        if (pin) target.ReuseTermWithPinning(term);
        else target.ReuseTerm(term, includeChildren);
        store.CommitAll();
        await ctx.ExecuteQueryAsync();
    }

    // ---- Plumbing ------------------------------------------------------------------------------------------

    /// <summary>The instance of a term in the term set it's shown under in the tree.</summary>
    private static Term GetTerm(TermStore store, TermNode node) => GetTerm(store, node.ToRef());

    /// <summary>The instance of a term in the term set the reference names.</summary>
    private static Term GetTerm(TermStore store, ItemRef item)
    {
        if (item.Kind != TermNodeKind.Term || item.TermSetId is not { } termSetId)
            throw new InvalidOperationException($"“{item.Name}” isn't a term in a term set.");
        return store.GetTermInTermSet(termSetId, item.Id);
    }

    private async Task<int> GetLanguageAsync(ConnectionProfile profile, ClientContext ctx, TermStore store)
    {
        if (_languages.TryGetValue(profile.Id, out var lcid)) return lcid;
        ctx.Load(store, s => s.DefaultLanguage);
        await ctx.ExecuteQueryAsync();
        return _languages[profile.Id] = store.DefaultLanguage;
    }

    private static TermStore GetStore(ClientContext ctx) =>
        TaxonomySession.GetTaxonomySession(ctx).GetDefaultSiteCollectionTermStore();

    private async Task<ClientContext> CreateContextAsync(ConnectionProfile profile, CancellationToken ct)
    {
        var token = await auth.GetAccessTokenAsync(profile, allowInteractive: true, ct: ct);
        var ctx = new ClientContext(profile.SiteUrl);
        ctx.ExecutingWebRequest += (_, e) => e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
        return ctx;
    }
}
