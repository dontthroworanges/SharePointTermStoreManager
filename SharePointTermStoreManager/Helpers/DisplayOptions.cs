using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Media;
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
/// Colourful icons are the filled glyphs, coloured by kind: amber groups, blue term sets, purple reused or pinned
/// terms and grey deprecated terms; plain terms keep the outline icon (see <see cref="IsColored"/>). Fluent's
/// full-colour icon set has no folder or tag icons, so filled glyphs in fixed colours stand in for it.
/// Templates show one of two FontIcons, so the outline one keeps the theme's text colour.
/// </summary>
public static class TermIcons
{
    private static readonly SolidColorBrush s_group = Brush(0xE8, 0xA3, 0x3D);
    private static readonly SolidColorBrush s_termSet = Brush(0x2B, 0x88, 0xD8);
    private static readonly SolidColorBrush s_linked = Brush(0x87, 0x64, 0xB8);
    private static readonly SolidColorBrush s_deprecated = Brush(0x9A, 0x98, 0x96);

    public static string Outline(TermNodeKind kind, bool deprecated, bool reused, bool pinned) => kind switch
    {
        TermNodeKind.Group => FluentIcons.Folder,
        TermNodeKind.TermSet => FluentIcons.GroupList,
        _ => deprecated ? FluentIcons.TagOff : reused || pinned ? FluentIcons.Link : FluentIcons.Tag,
    };

    public static string Filled(TermNodeKind kind, bool deprecated, bool reused, bool pinned) => kind switch
    {
        TermNodeKind.Group => FluentIcons.FolderFilled,
        TermNodeKind.TermSet => FluentIcons.GroupListFilled,
        _ => deprecated ? FluentIcons.TagOffFilled : reused || pinned ? FluentIcons.LinkFilled : FluentIcons.TagFilled,
    };

    public static SolidColorBrush Color(TermNodeKind kind, bool deprecated, bool reused, bool pinned) => kind switch
    {
        TermNodeKind.Group => s_group,
        TermNodeKind.TermSet => s_termSet,
        _ => deprecated ? s_deprecated : s_linked, // plain terms aren't coloured (see IsColored)
    };

    /// <summary>
    /// True when the item gets a coloured icon. Plain terms keep their outline icon in colourful mode too; only
    /// groups, term sets and reused, pinned or deprecated terms are coloured.
    /// </summary>
    public static bool IsColored(bool colorful, TermNodeKind kind, bool deprecated, bool reused, bool pinned) =>
        colorful && (kind != TermNodeKind.Term || deprecated || reused || pinned);

    public static Visibility OutlineVisibility(bool colorful, TermNodeKind kind, bool deprecated, bool reused, bool pinned) =>
        IsColored(colorful, kind, deprecated, reused, pinned) ? Visibility.Collapsed : Visibility.Visible;

    public static Visibility ColorVisibility(bool colorful, TermNodeKind kind, bool deprecated, bool reused, bool pinned) =>
        IsColored(colorful, kind, deprecated, reused, pinned) ? Visibility.Visible : Visibility.Collapsed;

    private static SolidColorBrush Brush(byte r, byte g, byte b) => new(Windows.UI.Color.FromArgb(0xFF, r, g, b));
}
