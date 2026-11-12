using System.Globalization;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.UI.Xaml;

namespace SharePointTermStoreManager.Models;

/// <summary>Kinds of text in a pretty-printed JSON line, for colouring.</summary>
public enum JsonTokenKind
{
    Punctuation,
    Key,
    String,
    Number,
    Literal,
}

/// <summary>One line of pretty-printed JSON: its nesting depth and its coloured pieces.</summary>
public sealed record JsonLine(int Depth, IReadOnlyList<(JsonTokenKind Kind, string Text)> Tokens);

/// <summary>One line of the operations log (see <see cref="Services.OperationLog"/>), parsed for the log page.</summary>
public sealed partial class LogEntry : ObservableObject
{
    private static readonly JsonSerializerOptions s_relaxed = new() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    private static readonly JsonSerializerOptions s_indented = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        WriteIndented = true,
    };

    private LogEntry(string raw, JsonElement root, string file)
    {
        Raw = raw;
        Root = root;
        File = file;

        Time = root.TryGetProperty("time", out var time) && time.TryGetDateTimeOffset(out var t) ? t : DateTimeOffset.MinValue;
        Site = GetString(root, "site");
        User = GetString(root, "user");
        Operation = GetString(root, "operation");
        Outcome = GetString(root, "outcome");
        Error = root.TryGetProperty("error", out var error) && error.ValueKind == JsonValueKind.String ? error.GetString() : null;

        // The item the operation was on, and where it went (a merge's target, a move's destination...).
        root.TryGetProperty("data", out var data);
        var item = Child(data, "item") ?? Child(data, "source");
        var other = Child(data, "target") ?? Child(data, "to");
        var name = item is { } i ? GetString(i, "Name") : "";
        if (other is { } o && GetString(o, "Name") is { Length: > 0 } otherName)
            name = name.Length > 0 ? $"{name} → {otherName}" : otherName;

        // An update: a rename shows old → new, and the changed fields follow the name.
        if (Child(data, "changes") is { } changes)
        {
            if (Child(changes, "Name") is { } rename && GetString(rename, "to") is { Length: > 0 } newName)
                name = $"{GetString(rename, "from")} → {newName}";
            ChangedFields = string.Join(", ", changes.EnumerateObject().Select(p => FieldLabel(p.Name)));
        }
        ItemName = name;
        ItemPath = item is { } p ? GetString(p, "Path") : "";

        var groups = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
        var text = new StringBuilder();
        Collect(root, groups, text);
        Groups = groups.ToList();
        text.Append(TimeText).Append('\n').Append(OperationText);
        SearchText = text.ToString();
    }

    /// <summary>The entry, or null when the line isn't a JSON object.</summary>
    public static LogEntry? TryParse(string line, string file)
    {
        try
        {
            using var doc = JsonDocument.Parse(line);
            return doc.RootElement.ValueKind == JsonValueKind.Object ? new LogEntry(line, doc.RootElement.Clone(), file) : null;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    /// <summary>The line as it is in the file.</summary>
    public string Raw { get; }

    public JsonElement Root { get; }

    /// <summary>The log file's name, e.g. "operations-2026-11.jsonl".</summary>
    public string File { get; }

    public DateTimeOffset Time { get; }
    public string Site { get; }
    public string User { get; }
    public string Operation { get; }
    public string Outcome { get; }
    public string? Error { get; }

    /// <summary>e.g. "Term 3 → Term 2" for a merge; empty when the entry names no item.</summary>
    public string ItemName { get; }

    public string ItemPath { get; }

    /// <summary>For an update, the fields it changed, e.g. "Synonyms, Description"; otherwise empty.</summary>
    public string ChangedFields { get; } = "";

    /// <summary>The term groups the entry mentions (from item paths and term group items).</summary>
    public IReadOnlyList<string> Groups { get; }

    /// <summary>Every value in the entry (one per line), plus the shown time and action, for the log's search box.</summary>
    public string SearchText { get; }

    /// <summary>Shown expanded (the JSON visible) or collapsed to its header.</summary>
    [ObservableProperty]
    private bool _isExpanded = true;

    // ---- Display -------------------------------------------------------------------------------------------

    public string TimeText => Time == DateTimeOffset.MinValue
        ? "Unknown time"
        : Time.ToLocalTime().ToString("ddd, MMM d, yyyy  h:mm:ss tt", CultureInfo.CurrentCulture);

    /// <summary>"Un-deprecate", "Merge"...</summary>
    public string OperationText => Capitalize(Operation.Length > 0 ? Operation : "Unknown action");

    public string OutcomeText => Capitalize(Outcome);

    /// <summary>"Merge: Term 3 → Term 2", "Update: Term 2 (Synonyms, Description)".</summary>
    public string Title
    {
        get
        {
            var title = ItemName.Length > 0 ? $"{OperationText}: {ItemName}" : OperationText;
            return ChangedFields.Length > 0 ? $"{title} ({ChangedFields})" : title;
        }
    }

    /// <summary>The item's path, or the site when there's no path.</summary>
    public string Subtitle => ItemPath.Length > 0 ? ItemPath : Site;

    public Visibility SucceededVisibility => Outcome == "succeeded" ? Visibility.Visible : Visibility.Collapsed;
    public Visibility FailedVisibility => Outcome == "failed" ? Visibility.Visible : Visibility.Collapsed;
    public Visibility OtherOutcomeVisibility => Outcome is "succeeded" or "failed" ? Visibility.Collapsed : Visibility.Visible;
    public Visibility ErrorVisibility => string.IsNullOrEmpty(Error) ? Visibility.Collapsed : Visibility.Visible;

    /// <summary>The entry as indented JSON, for copying.</summary>
    public string FormattedJson => JsonSerializer.Serialize(Root, s_indented);

    /// <summary>Several entries as an indented JSON array, for copying.</summary>
    public static string FormatAll(IEnumerable<LogEntry> entries) =>
        JsonSerializer.Serialize(entries.Select(e => e.Root).ToList(), s_indented);

    private IReadOnlyList<JsonLine>? _lines;

    /// <summary>The entry pretty-printed: one item per line, nested objects indented a level each.</summary>
    public IReadOnlyList<JsonLine> Lines => _lines ??= PrettyPrint(Root);

    // ---- Pretty printing -----------------------------------------------------------------------------------

    private static IReadOnlyList<JsonLine> PrettyPrint(JsonElement root)
    {
        var lines = new List<JsonLine>();
        var current = new List<(JsonTokenKind, string)>();
        int currentDepth = 0;

        void StartLine(int depth)
        {
            if (current.Count > 0) lines.Add(new JsonLine(currentDepth, current));
            current = new List<(JsonTokenKind, string)>();
            currentDepth = depth;
        }

        void Write(JsonElement value, int depth, string? key, bool comma)
        {
            StartLine(depth);
            if (key != null)
            {
                current.Add((JsonTokenKind.Key, JsonSerializer.Serialize(key, s_relaxed)));
                current.Add((JsonTokenKind.Punctuation, ": "));
            }

            switch (value.ValueKind)
            {
                case JsonValueKind.Object when value.EnumerateObject().Any():
                {
                    current.Add((JsonTokenKind.Punctuation, "{"));
                    var properties = value.EnumerateObject().ToList();
                    for (int i = 0; i < properties.Count; i++)
                        Write(properties[i].Value, depth + 1, properties[i].Name, i < properties.Count - 1);
                    StartLine(depth);
                    current.Add((JsonTokenKind.Punctuation, "}"));
                    break;
                }
                case JsonValueKind.Array when value.GetArrayLength() > 0:
                {
                    current.Add((JsonTokenKind.Punctuation, "["));
                    var items = value.EnumerateArray().ToList();
                    for (int i = 0; i < items.Count; i++) Write(items[i], depth + 1, null, i < items.Count - 1);
                    StartLine(depth);
                    current.Add((JsonTokenKind.Punctuation, "]"));
                    break;
                }
                case JsonValueKind.Object:
                    current.Add((JsonTokenKind.Punctuation, "{}"));
                    break;
                case JsonValueKind.Array:
                    current.Add((JsonTokenKind.Punctuation, "[]"));
                    break;
                case JsonValueKind.String:
                    current.Add((JsonTokenKind.String, JsonSerializer.Serialize(value.GetString(), s_relaxed)));
                    break;
                case JsonValueKind.Number:
                    current.Add((JsonTokenKind.Number, value.GetRawText()));
                    break;
                default: // true, false, null
                    current.Add((JsonTokenKind.Literal, value.GetRawText()));
                    break;
            }
            if (comma) current.Add((JsonTokenKind.Punctuation, ","));
        }

        Write(root, 0, null, false);
        StartLine(0);
        return lines;
    }

    // ---- Helpers -------------------------------------------------------------------------------------------

    private static JsonElement? Child(JsonElement parent, string name) =>
        parent.ValueKind == JsonValueKind.Object && parent.TryGetProperty(name, out var child) && child.ValueKind == JsonValueKind.Object
            ? child
            : null;

    private static string GetString(JsonElement parent, string name) =>
        parent.ValueKind == JsonValueKind.Object && parent.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? ""
            : "";

    /// <summary>Gathers term group names (from paths and group items) and all keys and values, for filtering.</summary>
    private static void Collect(JsonElement element, SortedSet<string> groups, StringBuilder text)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                if (GetString(element, "Path") is { Length: > 0 } path)
                    groups.Add(path.Split(" › ")[0].Trim());
                else if (GetString(element, "Kind") == "Group" && GetString(element, "Name") is { Length: > 0 } group)
                    groups.Add(group);
                // Values only: every entry has the same keys, so matching them would match everything.
                foreach (var property in element.EnumerateObject()) Collect(property.Value, groups, text);
                break;
            case JsonValueKind.Array:
                foreach (var item in element.EnumerateArray()) Collect(item, groups, text);
                break;
            case JsonValueKind.String:
                text.Append(element.GetString()).Append('\n');
                break;
            default:
                text.Append(element.GetRawText()).Append('\n');
                break;
        }
    }

    /// <summary>Field names as the editor labels them.</summary>
    private static string FieldLabel(string field) => field switch
    {
        "CustomProperties" => "Custom properties",
        "LocalCustomProperties" => "Local custom properties",
        "IsOpenForTermCreation" => "Open for term creation",
        "IsAvailableForTagging" => "Available for tagging",
        _ => field,
    };

    private static string Capitalize(string text) =>
        text.Length == 0 ? text : char.ToUpper(text[0], CultureInfo.CurrentCulture) + text[1..];
}
