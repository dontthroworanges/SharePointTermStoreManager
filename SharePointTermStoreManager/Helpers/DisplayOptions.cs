using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using SharePointTermStoreManager.Models;

namespace SharePointTermStoreManager.Helpers;

/// <summary>
/// Live display preferences that templates bind to with <c>{x:Bind h:DisplayOptions.Current.…, Mode=OneWay}</c>,
/// so a change in Settings updates the tree and lists straight away (as in SharePoint Search Explorer).
/// </summary>
public sealed partial class DisplayOptions : ObservableObject
{
    public static DisplayOptions Current { get; } = new();

    /// <summary>Colourful icons in the tree and lists instead of outline glyphs.</summary>
    [ObservableProperty]
    private bool _useColorfulIcons;
}

/// <summary>
/// The icon for a term group, term set or term. Outline icons are the regular glyphs in the text colour.
/// Colourful icons are full-colour SVGs in the style of Fluent's colour icon set (Assets/TermIcons): the
/// official Text Bullet List Square (term sets), Link (reused terms) and Pin (pinned terms) colour icons, plus a
/// folder (groups), tag (terms) and slashed grey tag (deprecated terms) drawn in the same style, because Fluent's
/// colour set has no folder or tag. Templates show either the outline FontIcon or the colour Image, so the
/// outline one keeps the theme's text colour.
/// </summary>
public static class TermIcons
{
    public static string Outline(TermNodeKind kind, bool deprecated, bool reused, bool pinned) => kind switch
    {
        TermNodeKind.Group => FluentIcons.Folder,
        TermNodeKind.TermSet => FluentIcons.GroupList,
        _ => deprecated ? FluentIcons.TagOff : reused || pinned ? FluentIcons.Link : FluentIcons.Tag,
    };

    /// <summary>The colour icon; deprecated wins over pinned, and pinned over reused.</summary>
    public static ImageSource ColorIcon(TermNodeKind kind, bool deprecated, bool reused, bool pinned) => Svg(kind switch
    {
        TermNodeKind.Group => "term-group",
        TermNodeKind.TermSet => "term-set",
        _ => deprecated ? "term-deprecated" : pinned ? "term-pinned" : reused ? "term-reused" : "term",
    });

    public static Visibility OutlineVisibility(bool colorful) => colorful ? Visibility.Collapsed : Visibility.Visible;

    public static Visibility ColorVisibility(bool colorful) => colorful ? Visibility.Visible : Visibility.Collapsed;

    /// <summary>One shared image source per icon (they're only created on the UI thread, by bindings).</summary>
    private static readonly Dictionary<string, SvgImageSource> s_icons = new();

    private static SvgImageSource Svg(string name)
    {
        if (!s_icons.TryGetValue(name, out var source))
            s_icons[name] = source = new SvgImageSource(new Uri($"ms-appx:///Assets/TermIcons/{name}.svg"));
        return source;
    }
}
