using System.Diagnostics;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using SharePointTermStoreManager.Models;

namespace SharePointTermStoreManager.Services;

/// <summary>
/// Appends a record of every change the app makes to the term store to a monthly JSON Lines file in
/// %LOCALAPPDATA%\SharePointTermStoreManager\Logs. Before a merge or delete the item's own details (labels,
/// description, properties, IDs) are written, so what was removed can be looked up or recreated by hand.
/// Child terms aren't included.
/// </summary>
public static class OperationLog
{
    public static string Folder { get; } = Path.Combine(SettingsStore.DataFolder, "Logs");

    private static readonly JsonSerializerOptions s_json = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, // keep non-ASCII labels readable
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() },
    };

    private static readonly object s_lock = new();

    /// <summary>Raised after an entry is written or the log is deleted (on the thread that did it).</summary>
    public static event EventHandler? Changed;

    /// <summary>The monthly log files, newest first.</summary>
    public static IReadOnlyList<string> Files() =>
        Directory.Exists(Folder)
            ? Directory.GetFiles(Folder, "operations-*.jsonl").OrderByDescending(Path.GetFileName, StringComparer.OrdinalIgnoreCase).ToList()
            : [];

    /// <summary>
    /// Reads every entry in every log file, newest first. Lines that aren't valid JSON are skipped and counted
    /// in <paramref name="skipped"/>.
    /// </summary>
    public static IReadOnlyList<LogEntry> ReadAll(out int skipped)
    {
        var entries = new List<LogEntry>();
        skipped = 0;
        foreach (var file in Files())
        {
            string[] lines;
            try
            {
                lock (s_lock)
                {
                    using var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var reader = new StreamReader(stream);
                    lines = reader.ReadToEnd().Split('\n');
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Debug.WriteLine($"Could not read {file}: {ex.Message}");
                continue;
            }

            foreach (var raw in lines)
            {
                var line = raw.Trim();
                if (line.Length == 0) continue;
                if (LogEntry.TryParse(line, Path.GetFileName(file)) is { } entry) entries.Add(entry);
                else skipped++;
            }
        }
        return entries.OrderByDescending(e => e.Time).ToList();
    }

    /// <summary>Deletes every log file (for Settings › Reset, when asked). Best effort.</summary>
    public static void DeleteAll()
    {
        try
        {
            lock (s_lock)
            {
                if (Directory.Exists(Folder)) Directory.Delete(Folder, recursive: true);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Debug.WriteLine($"Could not delete the operations log: {ex.Message}");
        }
        Changed?.Invoke(null, EventArgs.Empty);
    }

    /// <param name="operation">e.g. "merge", "delete", "move".</param>
    /// <param name="outcome">"started", "succeeded" or "failed".</param>
    public static void Write(ConnectionProfile profile, string operation, string outcome, object? data = null, Exception? error = null)
    {
        try
        {
            var entry = new
            {
                time = DateTimeOffset.Now,
                site = profile.SiteUrl,
                user = profile.UserName,
                operation,
                outcome,
                data,
                error = error?.Message,
            };
            var line = JsonSerializer.Serialize(entry, s_json);
            lock (s_lock)
            {
                Directory.CreateDirectory(Folder);
                File.AppendAllText(Path.Combine(Folder, $"operations-{DateTime.Now:yyyy-MM}.jsonl"), line + Environment.NewLine);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Could not write the operations log: {ex.Message}");
            return;
        }
        Changed?.Invoke(null, EventArgs.Empty);
    }
}
