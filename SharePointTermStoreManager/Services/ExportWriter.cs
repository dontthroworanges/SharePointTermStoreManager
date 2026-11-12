using System.Globalization;
using System.IO.Compression;
using System.Net;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml;
using SharePointTermStoreManager.Models;

namespace SharePointTermStoreManager.Services;

/// <summary>
/// Writes an <see cref="ExportStore"/> as pretty-printed JSON (nested like the tree), an Excel workbook or CSV
/// file (one row per term group, term set and term), or a single self-contained HTML page with a collapsible
/// tree, search and a details pane.
/// </summary>
public static class ExportWriter
{
    public static string Extension(ExportFormat format) => format switch
    {
        ExportFormat.Json => ".json",
        ExportFormat.Excel => ".xlsx",
        ExportFormat.Csv => ".csv",
        _ => ".html",
    };

    public static string FileTypeName(ExportFormat format) => format switch
    {
        ExportFormat.Json => "JSON",
        ExportFormat.Excel => "Excel workbook",
        ExportFormat.Csv => "CSV (comma separated)",
        _ => "Web page",
    };

    public static void Write(ExportStore store, ExportFormat format, string path)
    {
        switch (format)
        {
            case ExportFormat.Json: WriteJson(store, path); break;
            case ExportFormat.Excel: WriteExcel(store, path); break;
            case ExportFormat.Csv: WriteCsv(store, path); break;
            default: WriteHtml(store, path); break;
        }
    }

    // ---- JSON ----------------------------------------------------------------------------------------------

    private static readonly JsonSerializerOptions s_prettyJson = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, // keep non-ASCII labels readable
    };

    private static void WriteJson(ExportStore store, string path)
    {
        using var stream = File.Create(path);
        JsonSerializer.Serialize(stream, store, s_prettyJson);
    }

    // ---- Rows (Excel and CSV) ------------------------------------------------------------------------------

    /// <summary>A table column: its header, Excel width, and how to read its value from a row.</summary>
    private sealed record Column(string Header, double Width, Func<ExportRow, object?> Value);

    private static readonly Column[] s_columns =
    [
        new("Type", 11, r => r.Type),
        new("Term group", 22, r => r.TermGroup),
        new("Term set", 22, r => r.TermSet),
        new("Level", 8, r => r.Level),
        new("Name", 30, r => r.Name),
        new("Path", 50, r => r.Path),
        new("Description", 40, r => r.Description),
        new("Synonyms", 30, r => r.Synonyms),
        new("Other languages", 30, r => r.OtherLanguages),
        new("Available for tagging", 12, r => r.AvailableForTagging),
        new("Open for term creation", 12, r => r.OpenForTermCreation),
        new("Deprecated", 12, r => r.Deprecated),
        new("Reused", 10, r => r.Reused),
        new("Pinned", 10, r => r.Pinned),
        new("Source term", 10, r => r.SourceTerm),
        new("Owner", 22, r => r.Owner),
        new("Contact", 22, r => r.Contact),
        new("Custom properties", 30, r => r.CustomProperties),
        new("Local custom properties", 30, r => r.LocalCustomProperties),
        new("Child terms", 10, r => r.ChildTerms),
        new("ID", 38, r => r.Id),
        new("Parent ID", 38, r => r.ParentId),
        new("Term set ID", 38, r => r.TermSetId),
        new("Term group ID", 38, r => r.TermGroupId),
        new("Created", 18, r => r.Created),
        new("Last modified", 18, r => r.LastModified),
    ];

    /// <summary>The store as table rows, in tree order: each group, then each of its term sets and their terms.</summary>
    public static IEnumerable<ExportRow> Rows(ExportStore store)
    {
        foreach (var group in store.Groups)
        {
            yield return new ExportRow
            {
                Type = "Term group",
                TermGroup = group.Name,
                Name = group.Name,
                Path = group.Name,
                Description = group.Description,
                Id = group.Id,
                TermGroupId = group.Id,
                Created = group.Created,
                LastModified = group.LastModified,
            };
            foreach (var termSet in group.TermSets)
            {
                var setPath = $"{group.Name} › {termSet.Name}";
                yield return new ExportRow
                {
                    Type = "Term set",
                    TermGroup = group.Name,
                    TermSet = termSet.Name,
                    Name = termSet.Name,
                    Path = setPath,
                    Description = termSet.Description,
                    AvailableForTagging = termSet.IsAvailableForTagging,
                    OpenForTermCreation = termSet.IsOpenForTermCreation,
                    Owner = termSet.Owner,
                    Contact = termSet.Contact,
                    CustomProperties = Properties(termSet.CustomProperties),
                    ChildTerms = termSet.Terms.Count,
                    Id = termSet.Id,
                    ParentId = group.Id,
                    TermSetId = termSet.Id,
                    TermGroupId = group.Id,
                    Created = termSet.Created,
                    LastModified = termSet.LastModified,
                };
                foreach (var row in TermRows(termSet.Terms, 1, setPath, termSet.Id, group, termSet)) yield return row;
            }
        }
    }

    private static IEnumerable<ExportRow> TermRows(
        List<ExportTerm> terms, int level, string parentPath, Guid parentId, ExportGroup group, ExportTermSet termSet)
    {
        foreach (var term in terms)
        {
            var path = $"{parentPath} › {term.Name}";
            yield return new ExportRow
            {
                Type = "Term",
                TermGroup = group.Name,
                TermSet = termSet.Name,
                Level = level,
                Name = term.Name,
                Path = path,
                Description = term.Description,
                Synonyms = string.Join("; ", term.Synonyms),
                OtherLanguages = string.Join(" | ", (term.OtherLanguages ?? []).Select(l =>
                    l.Synonyms.Count > 0 ? $"{l.Language}: {l.Name} ({string.Join("; ", l.Synonyms)})" : $"{l.Language}: {l.Name}")),
                AvailableForTagging = term.IsAvailableForTagging,
                Deprecated = term.IsDeprecated,
                Reused = term.IsReused,
                Pinned = term.IsPinned,
                SourceTerm = term.IsSourceTerm,
                CustomProperties = Properties(term.CustomProperties),
                LocalCustomProperties = Properties(term.LocalCustomProperties),
                ChildTerms = term.Terms.Count,
                Id = term.Id,
                ParentId = parentId,
                TermSetId = termSet.Id,
                TermGroupId = group.Id,
                Created = term.Created,
                LastModified = term.LastModified,
            };
            foreach (var row in TermRows(term.Terms, level + 1, path, term.Id, group, termSet)) yield return row;
        }
    }

    /// <summary>"Key1=Value1; Key2=Value2".</summary>
    private static string Properties(Dictionary<string, string> properties) =>
        string.Join("; ", properties.Select(p => $"{p.Key}={p.Value}"));

    /// <summary>Local time as text, for CSV.</summary>
    private static string DateText(DateTime value) =>
        value == default ? "" : ToLocal(value).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

    private static DateTime ToLocal(DateTime value) =>
        value.Kind == DateTimeKind.Local ? value : DateTime.SpecifyKind(value, DateTimeKind.Utc).ToLocalTime();

    // ---- CSV -----------------------------------------------------------------------------------------------

    /// <summary>UTF-8 with a byte order mark (so Excel reads non-ASCII text correctly), comma separated, CRLF lines.</summary>
    private static void WriteCsv(ExportStore store, string path)
    {
        using var writer = new StreamWriter(path, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        writer.NewLine = "\r\n";
        writer.WriteLine(string.Join(",", s_columns.Select(c => Quote(c.Header))));
        foreach (var row in Rows(store))
            writer.WriteLine(string.Join(",", s_columns.Select(c => Quote(CsvValue(c.Value(row))))));
    }

    private static string CsvValue(object? value) => value switch
    {
        null => "",
        bool b => b ? "Yes" : "No",
        DateTime d => DateText(d),
        int i => i.ToString(CultureInfo.InvariantCulture),
        Guid g => g.ToString(),
        _ => value.ToString() ?? "",
    };

    private static string Quote(string value) =>
        value.IndexOfAny([',', '"', '\r', '\n']) >= 0 ? $"\"{value.Replace("\"", "\"\"")}\"" : value;

    // ---- Excel ---------------------------------------------------------------------------------------------

    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>
    /// An .xlsx workbook with one sheet holding an Excel table (filter buttons, banded rows, header frozen), so it
    /// opens ready to sort and filter. Written directly as Office Open XML, so no library is needed.
    /// </summary>
    private static void WriteExcel(ExportStore store, string path)
    {
        var rows = Rows(store).ToList();
        var strings = new Dictionary<string, int>(StringComparer.Ordinal);
        var stringList = new List<string>();
        int Shared(string text)
        {
            text = CleanXml(text);
            if (text.Length > 32767) text = text[..32767]; // Excel's limit for a cell
            if (!strings.TryGetValue(text, out var index))
            {
                index = stringList.Count;
                strings[text] = index;
                stringList.Add(text);
            }
            return index;
        }

        string lastColumn = ColumnName(s_columns.Length - 1);
        int lastRow = Math.Max(2, rows.Count + 1); // a table needs at least one data row
        string tableRef = $"A1:{lastColumn}{lastRow}";

        using var zip = new ZipArchive(File.Create(path), ZipArchiveMode.Create);

        WritePart(zip, "[Content_Types].xml", x =>
        {
            x.WriteStartElement("Types", "http://schemas.openxmlformats.org/package/2006/content-types");
            Default(x, "rels", "application/vnd.openxmlformats-package.relationships+xml");
            Default(x, "xml", "application/xml");
            Override(x, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            Override(x, "/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            Override(x, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            Override(x, "/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
            Override(x, "/xl/tables/table1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml");
            x.WriteEndElement();
        });

        WritePart(zip, "_rels/.rels", x => Relationships(x,
            ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml")));

        WritePart(zip, "xl/workbook.xml", x =>
        {
            x.WriteStartElement("workbook", MainNs);
            x.WriteAttributeString("xmlns", "r", null, RelNs);
            x.WriteStartElement("bookViews");
            x.WriteStartElement("workbookView");
            x.WriteEndElement();
            x.WriteEndElement();
            x.WriteStartElement("sheets");
            x.WriteStartElement("sheet");
            x.WriteAttributeString("name", "Term store");
            x.WriteAttributeString("sheetId", "1");
            x.WriteAttributeString("id", RelNs, "rId1");
            x.WriteEndElement();
            x.WriteEndElement();
            x.WriteEndElement();
        });

        WritePart(zip, "xl/_rels/workbook.xml.rels", x => Relationships(x,
            ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet1.xml"),
            ("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml"),
            ("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "sharedStrings.xml")));

        // Styles: 0 is the default; 1 shows a date and time.
        WritePart(zip, "xl/styles.xml", x =>
        {
            x.WriteStartElement("styleSheet", MainNs);
            x.WriteStartElement("numFmts");
            x.WriteAttributeString("count", "1");
            x.WriteStartElement("numFmt");
            x.WriteAttributeString("numFmtId", "164");
            x.WriteAttributeString("formatCode", "yyyy-mm-dd hh:mm");
            x.WriteEndElement();
            x.WriteEndElement();
            x.WriteRaw("<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>");
            x.WriteRaw("<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>");
            x.WriteRaw("<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>");
            x.WriteRaw("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
            x.WriteRaw("<cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
                       "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs>");
            x.WriteRaw("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
            x.WriteEndElement();
        });

        WritePart(zip, "xl/worksheets/sheet1.xml", x =>
        {
            x.WriteStartElement("worksheet", MainNs);
            x.WriteAttributeString("xmlns", "r", null, RelNs);

            // Keep the header row in view while scrolling.
            x.WriteRaw("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\">" +
                       "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>" +
                       "<selection pane=\"bottomLeft\" activeCell=\"A2\" sqref=\"A2\"/></sheetView></sheetViews>");
            x.WriteRaw("<sheetFormatPr defaultRowHeight=\"15\"/>");

            x.WriteStartElement("cols");
            for (int i = 0; i < s_columns.Length; i++)
            {
                x.WriteStartElement("col");
                x.WriteAttributeString("min", (i + 1).ToString(CultureInfo.InvariantCulture));
                x.WriteAttributeString("max", (i + 1).ToString(CultureInfo.InvariantCulture));
                x.WriteAttributeString("width", s_columns[i].Width.ToString(CultureInfo.InvariantCulture));
                x.WriteAttributeString("customWidth", "1");
                x.WriteEndElement();
            }
            x.WriteEndElement();

            x.WriteStartElement("sheetData");
            WriteRow(x, 1, s_columns.Select(c => (object?)c.Header).ToList(), Shared);
            for (int r = 0; r < rows.Count; r++)
                WriteRow(x, r + 2, s_columns.Select(c => c.Value(rows[r])).ToList(), Shared);
            x.WriteEndElement();

            x.WriteStartElement("tableParts");
            x.WriteAttributeString("count", "1");
            x.WriteStartElement("tablePart");
            x.WriteAttributeString("id", RelNs, "rId1");
            x.WriteEndElement();
            x.WriteEndElement();
            x.WriteEndElement();
        });

        WritePart(zip, "xl/worksheets/_rels/sheet1.xml.rels", x => Relationships(x,
            ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table", "../tables/table1.xml")));

        WritePart(zip, "xl/tables/table1.xml", x =>
        {
            x.WriteStartElement("table", MainNs);
            x.WriteAttributeString("id", "1");
            x.WriteAttributeString("name", "TermStore");
            x.WriteAttributeString("displayName", "TermStore");
            x.WriteAttributeString("ref", tableRef);
            x.WriteAttributeString("totalsRowShown", "0");
            x.WriteStartElement("autoFilter");
            x.WriteAttributeString("ref", tableRef);
            x.WriteEndElement();
            x.WriteStartElement("tableColumns");
            x.WriteAttributeString("count", s_columns.Length.ToString(CultureInfo.InvariantCulture));
            for (int i = 0; i < s_columns.Length; i++)
            {
                x.WriteStartElement("tableColumn");
                x.WriteAttributeString("id", (i + 1).ToString(CultureInfo.InvariantCulture));
                x.WriteAttributeString("name", s_columns[i].Header);
                x.WriteEndElement();
            }
            x.WriteEndElement();
            x.WriteStartElement("tableStyleInfo");
            x.WriteAttributeString("name", "TableStyleMedium2");
            x.WriteAttributeString("showFirstColumn", "0");
            x.WriteAttributeString("showLastColumn", "0");
            x.WriteAttributeString("showRowStripes", "1");
            x.WriteAttributeString("showColumnStripes", "0");
            x.WriteEndElement();
            x.WriteEndElement();
        });

        // Shared strings last: the sheet above added them.
        WritePart(zip, "xl/sharedStrings.xml", x =>
        {
            x.WriteStartElement("sst", MainNs);
            x.WriteAttributeString("count", stringList.Count.ToString(CultureInfo.InvariantCulture));
            x.WriteAttributeString("uniqueCount", stringList.Count.ToString(CultureInfo.InvariantCulture));
            foreach (var text in stringList)
            {
                x.WriteStartElement("si");
                x.WriteStartElement("t");
                if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1]) || text.Contains('\n')))
                    x.WriteAttributeString("xml", "space", null, "preserve");
                x.WriteString(text);
                x.WriteEndElement();
                x.WriteEndElement();
            }
            x.WriteEndElement();
        });
    }

    private static void WriteRow(XmlWriter x, int rowNumber, List<object?> values, Func<string, int> shared)
    {
        x.WriteStartElement("row");
        x.WriteAttributeString("r", rowNumber.ToString(CultureInfo.InvariantCulture));
        for (int i = 0; i < values.Count; i++)
        {
            var value = values[i];
            string? text = value switch
            {
                null => null,
                bool b => b ? "Yes" : "No",
                Guid g => g.ToString(),
                string s => s,
                _ => null,
            };
            if (value == null || text == "" || (value is DateTime d0 && d0 == default)) continue; // empty cell

            x.WriteStartElement("c");
            x.WriteAttributeString("r", $"{ColumnName(i)}{rowNumber}");
            switch (value)
            {
                case int number:
                    x.WriteElementString("v", MainNs, number.ToString(CultureInfo.InvariantCulture));
                    break;
                case DateTime date:
                    x.WriteAttributeString("s", "1");
                    x.WriteElementString("v", MainNs, ToLocal(date).ToOADate().ToString("R", CultureInfo.InvariantCulture));
                    break;
                default:
                    x.WriteAttributeString("t", "s");
                    x.WriteElementString("v", MainNs, shared(text!).ToString(CultureInfo.InvariantCulture));
                    break;
            }
            x.WriteEndElement();
        }
        x.WriteEndElement();
    }

    /// <summary>"A" for 0, "Z" for 25, "AA" for 26.</summary>
    private static string ColumnName(int index)
    {
        var name = "";
        for (index++; index > 0; index = (index - 1) / 26) name = (char)('A' + (index - 1) % 26) + name;
        return name;
    }

    /// <summary>Drops characters XML can't hold (control characters, unpaired surrogates).</summary>
    private static string CleanXml(string text)
    {
        if (text.All(c => XmlConvert.IsXmlChar(c))) return text;
        var builder = new StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++)
        {
            if (XmlConvert.IsXmlChar(text[i])) builder.Append(text[i]);
            else if (i + 1 < text.Length && XmlConvert.IsXmlSurrogatePair(text[i + 1], text[i]))
            {
                builder.Append(text[i]).Append(text[i + 1]);
                i++;
            }
        }
        return builder.ToString();
    }

    private static void WritePart(ZipArchive zip, string name, Action<XmlWriter> write)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var x = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false) });
        x.WriteStartDocument(standalone: true);
        write(x);
        x.WriteEndDocument();
    }

    private static void Default(XmlWriter x, string extension, string contentType)
    {
        x.WriteStartElement("Default");
        x.WriteAttributeString("Extension", extension);
        x.WriteAttributeString("ContentType", contentType);
        x.WriteEndElement();
    }

    private static void Override(XmlWriter x, string part, string contentType)
    {
        x.WriteStartElement("Override");
        x.WriteAttributeString("PartName", part);
        x.WriteAttributeString("ContentType", contentType);
        x.WriteEndElement();
    }

    private static void Relationships(XmlWriter x, params (string Id, string Type, string Target)[] relationships)
    {
        x.WriteStartElement("Relationships", PackageRelNs);
        foreach (var (id, type, target) in relationships)
        {
            x.WriteStartElement("Relationship");
            x.WriteAttributeString("Id", id);
            x.WriteAttributeString("Type", type);
            x.WriteAttributeString("Target", target);
            x.WriteEndElement();
        }
        x.WriteEndElement();
    }

    // ---- HTML ----------------------------------------------------------------------------------------------

    /// <summary>One tree node in the HTML page's data, with short property names to keep large exports small.</summary>
    private sealed class HtmlNode
    {
        [JsonPropertyName("p")] public int Parent { get; init; } = -1;
        [JsonPropertyName("t")] public string Type { get; init; } = "";
        [JsonPropertyName("n")] public string Name { get; init; } = "";
        [JsonPropertyName("id")] public string Id { get; init; } = "";
        [JsonPropertyName("d")] public string? Description { get; init; }
        [JsonPropertyName("sy")] public List<string>? Synonyms { get; init; }
        [JsonPropertyName("ol")] public List<string>? OtherLanguages { get; init; }
        [JsonPropertyName("f")] public List<string>? Flags { get; init; }
        [JsonPropertyName("ow")] public string? Owner { get; init; }
        [JsonPropertyName("co")] public string? Contact { get; init; }
        [JsonPropertyName("cp")] public Dictionary<string, string>? CustomProperties { get; init; }
        [JsonPropertyName("lp")] public Dictionary<string, string>? LocalCustomProperties { get; init; }
        [JsonPropertyName("cr")] public string? Created { get; init; }
        [JsonPropertyName("lm")] public string? LastModified { get; init; }
    }

    /// <summary>
    /// A single HTML file with no outside references: the data is embedded as JSON and a small script draws a
    /// collapsible tree (children drawn when first opened, so large stores stay quick), a search box and a
    /// details pane. Follows the browser's light or dark mode.
    /// </summary>
    private static void WriteHtml(ExportStore store, string path)
    {
        var nodes = new List<HtmlNode>();
        string? Date(DateTime value) => value == default ? null : ToLocal(value).ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        Dictionary<string, string>? Props(Dictionary<string, string> p) => p.Count > 0 ? p : null;
        string? Text(string value) => value.Length > 0 ? value : null;

        void AddTerms(List<ExportTerm> terms, int parent)
        {
            foreach (var term in terms)
            {
                var flags = new List<string>();
                if (term.IsDeprecated) flags.Add("Deprecated");
                if (term.IsPinned) flags.Add("Pinned");
                else if (term.IsReused) flags.Add("Reused");
                if (term.IsReused && term.IsSourceTerm) flags.Add("Source term");
                if (!term.IsAvailableForTagging) flags.Add("Not available for tagging");
                nodes.Add(new HtmlNode
                {
                    Parent = parent,
                    Type = "t",
                    Name = term.Name,
                    Id = term.Id.ToString(),
                    Description = Text(term.Description),
                    Synonyms = term.Synonyms.Count > 0 ? term.Synonyms : null,
                    OtherLanguages = term.OtherLanguages?.Select(l =>
                        l.Synonyms.Count > 0 ? $"{l.Language}: {l.Name} ({string.Join("; ", l.Synonyms)})" : $"{l.Language}: {l.Name}").ToList(),
                    Flags = flags.Count > 0 ? flags : null,
                    CustomProperties = Props(term.CustomProperties),
                    LocalCustomProperties = Props(term.LocalCustomProperties),
                    Created = Date(term.Created),
                    LastModified = Date(term.LastModified),
                });
                AddTerms(term.Terms, nodes.Count - 1);
            }
        }

        foreach (var group in store.Groups)
        {
            var groupFlags = new List<string>();
            if (group.IsSystemGroup) groupFlags.Add("System group");
            if (group.IsSiteCollectionGroup) groupFlags.Add("Site collection group");
            nodes.Add(new HtmlNode
            {
                Type = "g",
                Name = group.Name,
                Id = group.Id.ToString(),
                Description = Text(group.Description),
                Flags = groupFlags.Count > 0 ? groupFlags : null,
                Created = Date(group.Created),
                LastModified = Date(group.LastModified),
            });
            int groupIndex = nodes.Count - 1;
            foreach (var termSet in group.TermSets)
            {
                var setFlags = new List<string>();
                if (termSet.IsOpenForTermCreation) setFlags.Add("Open for term creation");
                if (!termSet.IsAvailableForTagging) setFlags.Add("Not available for tagging");
                if (termSet.HasCustomSortOrder) setFlags.Add("Custom sort order");
                nodes.Add(new HtmlNode
                {
                    Parent = groupIndex,
                    Type = "s",
                    Name = termSet.Name,
                    Id = termSet.Id.ToString(),
                    Description = Text(termSet.Description),
                    Flags = setFlags.Count > 0 ? setFlags : null,
                    Owner = Text(termSet.Owner),
                    Contact = Text(termSet.Contact),
                    CustomProperties = Props(termSet.CustomProperties),
                    Created = Date(termSet.Created),
                    LastModified = Date(termSet.LastModified),
                });
                AddTerms(termSet.Terms, nodes.Count - 1);
            }
        }

        var data = new
        {
            store = store.Name,
            site = store.Site,
            exported = store.Exported.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture),
            language = TermStoreService.LanguageName(store.DefaultLanguage),
            nodes,
        };
        // The default encoder escapes <, > and &, so the JSON can't end the <script> element it's in.
        var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull });

        var host = Uri.TryCreate(store.Site, UriKind.Absolute, out var uri) ? uri.Host : store.Site;
        var title = WebUtility.HtmlEncode($"Term store: {host}");
        var html = HtmlTemplate.Page.Replace("{{TITLE}}", title).Replace("{{DATA}}", json);
        File.WriteAllText(path, html, new UTF8Encoding(false));
    }
}
