using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.Foundation;

namespace SharePointTermStoreManager.Views;

/// <summary>
/// Lays out a breadcrumb in one row, the way SharePoint shows an item's path: the current item keeps its full
/// name where it can, and its ancestors share the rest of the width, so short names stay whole and long ones are
/// shortened with "…" equally. When even shortened ancestors don't fit, the leading ones are hidden behind an
/// overflow item ("…", a menu of the hidden ones).
/// </summary>
/// <remarks>
/// Children, in order: the overflow item, then one per ancestor (root first), then the current item. Ancestor and
/// overflow items include their trailing separator and must shorten their text (e.g. TextTrimming) when measured
/// narrower than they'd like.
/// </remarks>
public sealed partial class BreadcrumbPanel : Panel
{
    /// <summary>Ancestors aren't shortened below this (or their natural width, if less) before some are hidden.</summary>
    private const double AncestorMinWidth = 64;

    /// <summary>The current item isn't shortened below this (or its natural width) before ancestors are hidden.</summary>
    private const double CurrentMinWidth = 120;

    private double[] _widths = [];

    /// <summary>How many leading ancestors are hidden behind the overflow item (0: none, and it's hidden).</summary>
    public int HiddenCount { get; private set; }

    protected override Size MeasureOverride(Size availableSize)
    {
        int count = Children.Count;
        _widths = new double[count];
        if (count < 2)
        {
            HiddenCount = 0;
            return new Size(0, 0);
        }

        // Natural sizes.
        var natural = new double[count];
        for (int i = 0; i < count; i++)
        {
            Children[i].Measure(new Size(double.PositiveInfinity, availableSize.Height));
            natural[i] = Children[i].DesiredSize.Width;
        }

        int first = 1, last = count - 1, ancestors = last - first;
        double width = availableSize.Width;
        double overflowWidth = natural[0];
        double current = natural[last];

        // Hide as few leading ancestors as possible: the rest, at their minimum, plus the current item at its
        // minimum, must fit.
        int hidden = 0;
        if (!double.IsInfinity(width))
        {
            for (; hidden < ancestors; hidden++)
            {
                double needed = (hidden > 0 ? overflowWidth : 0) + Math.Min(current, CurrentMinWidth);
                for (int i = first + hidden; i < last; i++) needed += Math.Min(natural[i], AncestorMinWidth);
                if (needed <= width) break;
            }
        }
        HiddenCount = hidden;

        double overflow = hidden > 0 ? overflowWidth : 0;
        double ancestorMin = 0, ancestorNatural = 0;
        for (int i = first + hidden; i < last; i++)
        {
            ancestorMin += Math.Min(natural[i], AncestorMinWidth);
            ancestorNatural += natural[i];
        }

        // The current item takes what it needs, leaving the ancestors at least their minimum; the ancestors share
        // the rest, each capped at the same width (so only the longest are shortened), never below its minimum.
        double currentWidth = double.IsInfinity(width)
            ? current
            : Math.Max(0, Math.Min(current, width - overflow - ancestorMin));
        double room = double.IsInfinity(width) ? ancestorNatural : Math.Max(ancestorMin, width - overflow - currentWidth);
        double cap = CapFor(natural, first + hidden, last, room);

        _widths[0] = overflow;
        for (int i = first; i < last; i++)
            _widths[i] = i < first + hidden ? 0 : Math.Clamp(cap, Math.Min(natural[i], AncestorMinWidth), natural[i]);
        _widths[last] = currentWidth;

        double total = 0, height = 0;
        for (int i = 0; i < count; i++)
        {
            Children[i].Measure(new Size(_widths[i], availableSize.Height));
            if (_widths[i] <= 0) continue;
            total += _widths[i];
            height = Math.Max(height, Children[i].DesiredSize.Height);
        }
        return new Size(total, height);
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        double x = 0;
        for (int i = 0; i < Children.Count; i++)
        {
            double w = i < _widths.Length ? _widths[i] : 0;
            Children[i].Arrange(w > 0 ? new Rect(x, 0, w, finalSize.Height) : new Rect(0, 0, 0, 0));
            x += w;
        }
        return finalSize;
    }

    /// <summary>The width cap that makes the ancestors from <paramref name="from"/> (inclusive) to <paramref name="to"/> fill <paramref name="room"/>.</summary>
    private static double CapFor(double[] natural, int from, int to, double room)
    {
        double low = 0, high = 0;
        for (int i = from; i < to; i++) high = Math.Max(high, natural[i]);
        for (int step = 0; step < 30; step++)
        {
            double mid = (low + high) / 2, used = 0;
            for (int i = from; i < to; i++) used += Math.Clamp(mid, Math.Min(natural[i], AncestorMinWidth), natural[i]);
            if (used > room) high = mid; else low = mid;
        }
        return low;
    }
}
