using System.Text.Json.Serialization;
using CommunityToolkit.Mvvm.ComponentModel;

namespace SharePointTermStoreManager.Models;

/// <summary>
/// A saved connection: which SharePoint Online site to open the term store from, and how to sign in to it.
/// Any site in the tenant reaches the tenant-wide term store; term groups scoped to one site collection
/// only appear when the profile points at that site.
/// </summary>
public partial class ConnectionProfile : ObservableObject
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    [ObservableProperty]
    private string _name = "";

    [ObservableProperty]
    private string _siteUrl = "";

    /// <summary>Optional tenant ID or domain. Derived from the site URL when empty.</summary>
    [ObservableProperty]
    private string _tenantId = "";

    /// <summary>Optional Microsoft Entra app registration (public client). Empty uses the built-in client.</summary>
    [ObservableProperty]
    private string _clientId = "";

    /// <summary>MSAL home account identifier of the last account used with this profile.</summary>
    [ObservableProperty]
    private string _accountId = "";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(AccountDisplay))]
    private string _userName = "";

    /// <summary>True once a token has been confirmed this session. Not persisted.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(AccountDisplay))]
    [property: JsonIgnore]
    private bool _isSignedIn;

    [JsonIgnore]
    public string AccountDisplay =>
        IsSignedIn && !string.IsNullOrEmpty(UserName) ? $"Signed in as {UserName}" : "Not signed in";

    [JsonIgnore]
    public string EffectiveName => string.IsNullOrWhiteSpace(Name) ? SiteUrl : Name;
}
