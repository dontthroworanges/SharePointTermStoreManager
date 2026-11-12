using System.Text.RegularExpressions;

namespace SharePointTermStoreManager.Helpers;

/// <summary>
/// The term store's rules for names of groups, term sets and terms, and for labels: at most 255 characters, none
/// of <c>; " &lt; &gt; |</c> or tab, and an ampersand is stored as the full-width <c>＆</c> (U+FF06), as the
/// SharePoint term store page does.
/// </summary>
public static partial class TermNames
{
    public const int MaxNameLength = 255;
    public const int MaxDescriptionLength = 1000;

    private static readonly char[] s_invalid = [';', '"', '<', '>', '|', '\t'];

    /// <summary>Trims, collapses runs of white space to one space and replaces &amp; with ＆.</summary>
    public static string Normalize(string? name) =>
        WhiteSpace().Replace((name ?? "").Trim(), " ").Replace('&', '＆');

    /// <summary>An error message for a normalized name, or null when it's valid. <paramref name="what"/> is e.g. "name".</summary>
    public static string? Validate(string name, string what)
    {
        if (name.Length == 0) return $"Enter a {what}.";
        if (name.Length > MaxNameLength) return $"The {what} “{Shorten(name)}” is longer than {MaxNameLength} characters.";
        if (name.IndexOfAny(s_invalid) >= 0) return $"The {what} “{Shorten(name)}” can't contain ; \" < > or |.";
        return null;
    }

    public static string? ValidateDescription(string description) =>
        description.Length > MaxDescriptionLength ? $"The description is longer than {MaxDescriptionLength} characters." : null;

    private static string Shorten(string s) => s.Length <= 40 ? s : s[..40] + "…";

    [GeneratedRegex(@"\s+")]
    private static partial Regex WhiteSpace();
}
