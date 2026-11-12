using Microsoft.UI.Text;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Documents;
using Microsoft.UI.Xaml.Media;
using SharePointTermStoreManager.Models;
using Windows.UI;

namespace SharePointTermStoreManager.Views;

/// <summary>
/// A log entry's JSON, pretty-printed like Edge's JSON viewer: one value per line, each nesting level indented
/// with a faint guide line down its left side, and keys, strings, numbers and true/false/null in their own
/// colours. A long line wraps within its own indent, so the levels stay easy to follow.
/// </summary>
public sealed partial class JsonView : UserControl
{
    public static readonly DependencyProperty EntryProperty = DependencyProperty.Register(
        nameof(Entry), typeof(LogEntry), typeof(JsonView), new PropertyMetadata(null, (d, _) => ((JsonView)d).Build()));

    /// <summary>Width of one indent level.</summary>
    private const double IndentWidth = 20;

    private static readonly FontFamily s_mono = new("Cascadia Mono, Consolas, Courier New");

    private readonly StackPanel _lines = new();

    public JsonView()
    {
        Content = _lines;
        ActualThemeChanged += (_, _) => Build();
    }

    public LogEntry? Entry
    {
        get => (LogEntry?)GetValue(EntryProperty);
        set => SetValue(EntryProperty, value);
    }

    private void Build()
    {
        _lines.Children.Clear();
        if (Entry is not { } entry) return;

        var palette = ActualTheme == ElementTheme.Light ? Palette.Light : Palette.Dark;
        foreach (var line in entry.Lines)
        {
            var row = new Grid();
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            if (line.Depth > 0)
            {
                // One guide per level, under the opening bracket of the level above.
                var guides = new StackPanel { Orientation = Orientation.Horizontal };
                for (int i = 0; i < line.Depth; i++)
                {
                    guides.Children.Add(new Border
                    {
                        Width = IndentWidth - 4,
                        Margin = new Thickness(4, 0, 0, 0),
                        BorderThickness = new Thickness(1, 0, 0, 0),
                        BorderBrush = palette.Guide,
                    });
                }
                row.Children.Add(guides);
            }

            var text = new TextBlock
            {
                FontFamily = s_mono,
                FontSize = 13,
                TextWrapping = TextWrapping.Wrap,
                IsTextSelectionEnabled = true,
                Padding = new Thickness(0, 1, 0, 1),
            };
            foreach (var (kind, value) in line.Tokens)
            {
                text.Inlines.Add(new Run
                {
                    Text = value,
                    Foreground = palette.For(kind),
                    FontWeight = kind == JsonTokenKind.Key ? FontWeights.SemiBold : FontWeights.Normal,
                });
            }
            Grid.SetColumn(text, 1);
            row.Children.Add(text);
            _lines.Children.Add(row);
        }
    }

    /// <summary>Token colours for the light and dark themes (as in Visual Studio Code's default themes).</summary>
    private sealed record Palette(Brush Punctuation, Brush Key, Brush String, Brush Number, Brush Literal, Brush Guide)
    {
        public static readonly Palette Light = new(
            Solid(0xFF, 0x5F, 0x5F, 0x5F), Solid(0xFF, 0x04, 0x51, 0xA5), Solid(0xFF, 0xA3, 0x15, 0x15),
            Solid(0xFF, 0x09, 0x86, 0x58), Solid(0xFF, 0x00, 0x00, 0xFF), Solid(0x24, 0x00, 0x00, 0x00));

        public static readonly Palette Dark = new(
            Solid(0xFF, 0xB4, 0xB4, 0xB4), Solid(0xFF, 0x9C, 0xDC, 0xFE), Solid(0xFF, 0xCE, 0x91, 0x78),
            Solid(0xFF, 0xB5, 0xCE, 0xA8), Solid(0xFF, 0x56, 0x9C, 0xD6), Solid(0x30, 0xFF, 0xFF, 0xFF));

        public Brush For(JsonTokenKind kind) => kind switch
        {
            JsonTokenKind.Key => Key,
            JsonTokenKind.String => String,
            JsonTokenKind.Number => Number,
            JsonTokenKind.Literal => Literal,
            _ => Punctuation,
        };

        private static SolidColorBrush Solid(byte a, byte r, byte g, byte b) => new(Color.FromArgb(a, r, g, b));
    }
}
