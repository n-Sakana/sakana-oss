using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

public sealed class ExtractMdTextResult
{
    public string Content;
    public string EncodingName;
}

public sealed class ExtractMdOpenXmlResult
{
    public string Content;
    public string Notes;
}

public static class ExtractMdNativeHelper
{
    private const long MaxXmlEntryBytes = 64L * 1024L * 1024L;
    private const long MaxOpenXmlBytes = 256L * 1024L * 1024L;
    private const int MaxOpenXmlEntries = 10000;
    private const int MaxExcelColumn = 16384;

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    public static int GetWindowProcessId(IntPtr window)
    {
        if (window == IntPtr.Zero) return 0;
        uint processId;
        GetWindowThreadProcessId(window, out processId);
        return processId > Int32.MaxValue ? 0 : (int)processId;
    }

    public static ExtractMdTextResult ReadTextFile(string path)
    {
        byte[] bytes = File.ReadAllBytes(path);
        ExtractMdTextResult result = new ExtractMdTextResult();

        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            result.Content = Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
            result.EncodingName = "UTF-8 BOM";
            return result;
        }

        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
        {
            result.Content = Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
            result.EncodingName = "UTF-16 LE";
            return result;
        }

        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
        {
            result.Content = Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
            result.EncodingName = "UTF-16 BE";
            return result;
        }

        try
        {
            UTF8Encoding utf8 = new UTF8Encoding(false, true);
            result.Content = utf8.GetString(bytes);
            result.EncodingName = "UTF-8";
            return result;
        }
        catch
        {
        }

        try
        {
            result.Content = Encoding.GetEncoding(932).GetString(bytes);
            result.EncodingName = "CP932";
            return result;
        }
        catch
        {
            result.Content = Encoding.Default.GetString(bytes);
            result.EncodingName = "Default";
            return result;
        }
    }

    public static ExtractMdOpenXmlResult ReadWordOpenXml(string path)
    {
        StringBuilder builder = new StringBuilder();
        int partCount = 0;

        using (FileStream stream = File.OpenRead(path))
        using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
        {
            ValidateOpenXmlArchive(archive);
            string[] entries = new string[] {
                "word/document.xml",
                "word/footnotes.xml",
                "word/endnotes.xml",
                "word/comments.xml"
            };

            for (int i = 0; i < entries.Length; i++)
            {
                ZipArchiveEntry entry = GetEntry(archive, entries[i]);
                if (entry == null)
                {
                    continue;
                }

                AppendWordPart(archive, entry.FullName, GetWordPartLabel(entry.FullName), builder);
                partCount++;
            }

            partCount += AppendMatchingWordParts(archive, @"^word/headers/header[0-9]+\.xml$", "Header", builder);
            partCount += AppendMatchingWordParts(archive, @"^word/footers/footer[0-9]+\.xml$", "Footer", builder);
        }

        ExtractMdOpenXmlResult result = new ExtractMdOpenXmlResult();
        result.Content = TrimEndLines(builder.ToString());
        result.Notes = "OpenXML parts: " + partCount.ToString();
        if (String.IsNullOrWhiteSpace(result.Content))
        {
            result.Notes += "; No text extracted";
        }
        return result;
    }

    public static ExtractMdOpenXmlResult ReadPowerPointOpenXml(string path)
    {
        StringBuilder builder = new StringBuilder();
        int slideCount = 0;
        int noteCount = 0;

        using (FileStream stream = File.OpenRead(path))
        using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
        {
            ValidateOpenXmlArchive(archive);
            List<ZipArchiveEntry> slides = GetMatchingEntries(archive, @"^ppt/slides/slide[0-9]+\.xml$");
            slides.Sort(CompareEntryNumber);

            foreach (ZipArchiveEntry entry in slides)
            {
                int number = GetLastNumber(entry.FullName);
                if (builder.Length > 0)
                {
                    builder.AppendLine();
                }
                builder.AppendLine("## Slide " + number.ToString());
                AppendPresentationXml(entry, builder);
                slideCount++;
            }

            List<ZipArchiveEntry> notes = GetMatchingEntries(archive, @"^ppt/notesSlides/notesSlide[0-9]+\.xml$");
            notes.Sort(CompareEntryNumber);

            foreach (ZipArchiveEntry entry in notes)
            {
                int number = GetLastNumber(entry.FullName);
                if (builder.Length > 0)
                {
                    builder.AppendLine();
                }
                builder.AppendLine("## Notes " + number.ToString());
                AppendPresentationXml(entry, builder);
                noteCount++;
            }
        }

        ExtractMdOpenXmlResult result = new ExtractMdOpenXmlResult();
        result.Content = TrimEndLines(builder.ToString());
        result.Notes = "Slides: " + slideCount.ToString() + "; Notes: " + noteCount.ToString();
        if (String.IsNullOrWhiteSpace(result.Content))
        {
            result.Notes += "; No text extracted";
        }
        return result;
    }

    public static ExtractMdOpenXmlResult ReadExcelOpenXml(string path)
    {
        StringBuilder builder = new StringBuilder();
        int sheetCount = 0;

        using (FileStream stream = File.OpenRead(path))
        using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
        {
            ValidateOpenXmlArchive(archive);
            List<string> sharedStrings = LoadSharedStrings(archive);
            List<SheetInfo> sheets = LoadWorkbookSheets(archive);

            if (sheets.Count == 0)
            {
                List<ZipArchiveEntry> sheetEntries = GetMatchingEntries(archive, @"^xl/worksheets/sheet[0-9]+\.xml$");
                sheetEntries.Sort(CompareEntryNumber);
                foreach (ZipArchiveEntry entry in sheetEntries)
                {
                    sheets.Add(new SheetInfo(entry.FullName, Path.GetFileNameWithoutExtension(entry.FullName)));
                }
            }

            foreach (SheetInfo sheet in sheets)
            {
                ZipArchiveEntry entry = GetEntry(archive, sheet.Path);
                if (entry == null)
                {
                    continue;
                }

                if (builder.Length > 0)
                {
                    builder.AppendLine();
                }
                builder.AppendLine("## Sheet: " + sheet.Name);
                AppendWorksheetXml(entry, sharedStrings, builder);
                sheetCount++;
            }
        }

        ExtractMdOpenXmlResult result = new ExtractMdOpenXmlResult();
        result.Content = TrimEndLines(builder.ToString());
        result.Notes = "Sheets: " + sheetCount.ToString();
        if (String.IsNullOrWhiteSpace(result.Content))
        {
            result.Notes += "; No text extracted";
        }
        return result;
    }

    private sealed class SheetInfo
    {
        public string Path;
        public string Name;

        public SheetInfo(string path, string name)
        {
            Path = NormalizePackagePath(path);
            Name = name;
        }
    }

    private static void AppendWordPart(ZipArchive archive, string entryName, string label, StringBuilder builder)
    {
        ZipArchiveEntry entry = GetEntry(archive, entryName);
        if (entry == null)
        {
            return;
        }

        if (builder.Length > 0)
        {
            builder.AppendLine();
        }
        builder.AppendLine("## " + label);
        XmlDocument document = LoadEntryXml(entry);
        AppendWordNode(document.DocumentElement, builder);
    }

    private static int AppendMatchingWordParts(ZipArchive archive, string pattern, string labelPrefix, StringBuilder builder)
    {
        int count = 0;
        List<ZipArchiveEntry> entries = GetMatchingEntries(archive, pattern);
        entries.Sort(CompareEntryNumber);
        foreach (ZipArchiveEntry entry in entries)
        {
            string label = labelPrefix + " " + GetLastNumber(entry.FullName).ToString();
            AppendWordPart(archive, entry.FullName, label, builder);
            count++;
        }
        return count;
    }

    private static string GetWordPartLabel(string entryName)
    {
        if (entryName.EndsWith("/document.xml", StringComparison.OrdinalIgnoreCase)) return "Document";
        if (entryName.EndsWith("/footnotes.xml", StringComparison.OrdinalIgnoreCase)) return "Footnotes";
        if (entryName.EndsWith("/endnotes.xml", StringComparison.OrdinalIgnoreCase)) return "Endnotes";
        if (entryName.EndsWith("/comments.xml", StringComparison.OrdinalIgnoreCase)) return "Comments";
        return entryName;
    }

    private static void AppendWordNode(XmlNode node, StringBuilder builder)
    {
        if (node == null)
        {
            return;
        }

        string localName = node.LocalName;
        if (localName == "t" || localName == "instrText" || localName == "delText")
        {
            builder.Append(node.InnerText);
            return;
        }
        if (localName == "tab")
        {
            builder.Append('\t');
            return;
        }
        if (localName == "br" || localName == "cr")
        {
            builder.AppendLine();
            return;
        }

        foreach (XmlNode child in node.ChildNodes)
        {
            AppendWordNode(child, builder);
        }

        if (localName == "p" || localName == "tr")
        {
            AppendLineIfNeeded(builder);
        }
        else if (localName == "tc")
        {
            builder.Append('\t');
        }
    }

    private static void AppendPresentationXml(ZipArchiveEntry entry, StringBuilder builder)
    {
        XmlDocument document = LoadEntryXml(entry);
        AppendPresentationNode(document.DocumentElement, builder);
    }

    private static void AppendPresentationNode(XmlNode node, StringBuilder builder)
    {
        if (node == null)
        {
            return;
        }

        string localName = node.LocalName;
        if (localName == "t")
        {
            builder.Append(node.InnerText);
            return;
        }
        if (localName == "br")
        {
            builder.AppendLine();
            return;
        }

        foreach (XmlNode child in node.ChildNodes)
        {
            AppendPresentationNode(child, builder);
        }

        if (localName == "p")
        {
            AppendLineIfNeeded(builder);
        }
    }

    private static List<string> LoadSharedStrings(ZipArchive archive)
    {
        List<string> values = new List<string>();
        ZipArchiveEntry entry = GetEntry(archive, "xl/sharedStrings.xml");
        if (entry == null)
        {
            return values;
        }

        XmlDocument document = LoadEntryXml(entry);
        List<XmlNode> items = new List<XmlNode>();
        CollectNodesByLocalName(document.DocumentElement, "si", items);
        foreach (XmlNode item in items)
        {
            values.Add(CollapseCellText(CollectTextFromNode(item)));
        }
        return values;
    }

    private static List<SheetInfo> LoadWorkbookSheets(ZipArchive archive)
    {
        List<SheetInfo> sheets = new List<SheetInfo>();
        ZipArchiveEntry workbookEntry = GetEntry(archive, "xl/workbook.xml");
        if (workbookEntry == null)
        {
            return sheets;
        }

        Dictionary<string, string> relationships = LoadWorkbookRelationships(archive);
        XmlDocument document = LoadEntryXml(workbookEntry);
        List<XmlNode> sheetNodes = new List<XmlNode>();
        CollectNodesByLocalName(document.DocumentElement, "sheet", sheetNodes);

        foreach (XmlNode sheetNode in sheetNodes)
        {
            string name = GetAttributeValue(sheetNode, "name");
            string relationshipId = GetAttributeValue(sheetNode, "id");
            if (String.IsNullOrEmpty(name))
            {
                name = "Sheet " + (sheets.Count + 1).ToString();
            }

            string target;
            if (!String.IsNullOrEmpty(relationshipId) && relationships.TryGetValue(relationshipId, out target))
            {
                sheets.Add(new SheetInfo(target, name));
            }
        }

        return sheets;
    }

    private static Dictionary<string, string> LoadWorkbookRelationships(ZipArchive archive)
    {
        Dictionary<string, string> map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        ZipArchiveEntry relsEntry = GetEntry(archive, "xl/_rels/workbook.xml.rels");
        if (relsEntry == null)
        {
            return map;
        }

        XmlDocument document = LoadEntryXml(relsEntry);
        List<XmlNode> relationships = new List<XmlNode>();
        CollectNodesByLocalName(document.DocumentElement, "Relationship", relationships);
        foreach (XmlNode relationship in relationships)
        {
            string id = GetAttributeValue(relationship, "Id");
            string target = GetAttributeValue(relationship, "Target");
            if (String.IsNullOrEmpty(id) || String.IsNullOrEmpty(target))
            {
                continue;
            }
            map[id] = NormalizeWorkbookTarget(target);
        }
        return map;
    }

    private static void AppendWorksheetXml(ZipArchiveEntry entry, List<string> sharedStrings, StringBuilder builder)
    {
        XmlDocument document = LoadEntryXml(entry);
        List<XmlNode> rows = new List<XmlNode>();
        CollectNodesByLocalName(document.DocumentElement, "row", rows);

        foreach (XmlNode row in rows)
        {
            SortedDictionary<int, string> cells = new SortedDictionary<int, string>();
            int fallbackColumn = 1;
            foreach (XmlNode child in row.ChildNodes)
            {
                if (child.LocalName != "c")
                {
                    continue;
                }

                string reference = GetAttributeValue(child, "r");
                int columnIndex = GetColumnIndex(reference);
                if (columnIndex <= 0)
                {
                    columnIndex = fallbackColumn;
                }

                cells[columnIndex] = ReadCellValue(child, sharedStrings);
                fallbackColumn = columnIndex + 1;
            }

            if (cells.Count == 0)
            {
                builder.AppendLine();
                continue;
            }

            int maxColumn = 0;
            foreach (int column in cells.Keys)
            {
                if (column > maxColumn) maxColumn = column;
            }

            string[] values = new string[maxColumn];
            for (int i = 0; i < values.Length; i++) values[i] = "";
            foreach (KeyValuePair<int, string> cell in cells)
            {
                if (cell.Key > 0 && cell.Key <= values.Length)
                {
                    values[cell.Key - 1] = CollapseCellText(cell.Value);
                }
            }
            builder.AppendLine(String.Join("\t", values));
        }
    }

    private static string ReadCellValue(XmlNode cell, List<string> sharedStrings)
    {
        string type = GetAttributeValue(cell, "t");
        if (String.Equals(type, "inlineStr", StringComparison.OrdinalIgnoreCase))
        {
            return CollectTextFromNode(cell);
        }

        XmlNode valueNode = FindFirstDescendantByLocalName(cell, "v");
        string rawValue = valueNode == null ? "" : valueNode.InnerText;

        if (String.Equals(type, "s", StringComparison.OrdinalIgnoreCase))
        {
            int index;
            if (Int32.TryParse(rawValue, out index) && index >= 0 && index < sharedStrings.Count)
            {
                return sharedStrings[index];
            }
            return rawValue;
        }

        if (String.Equals(type, "str", StringComparison.OrdinalIgnoreCase))
        {
            return rawValue;
        }

        return rawValue;
    }

    private static string CollectTextFromNode(XmlNode node)
    {
        StringBuilder builder = new StringBuilder();
        CollectTextFromNode(node, builder);
        return builder.ToString();
    }

    private static void CollectTextFromNode(XmlNode node, StringBuilder builder)
    {
        if (node == null)
        {
            return;
        }
        if (node.LocalName == "t")
        {
            builder.Append(node.InnerText);
            return;
        }
        foreach (XmlNode child in node.ChildNodes)
        {
            CollectTextFromNode(child, builder);
        }
    }

    private static void CollectNodesByLocalName(XmlNode node, string localName, List<XmlNode> results)
    {
        if (node == null)
        {
            return;
        }
        if (node.LocalName == localName)
        {
            results.Add(node);
        }
        foreach (XmlNode child in node.ChildNodes)
        {
            CollectNodesByLocalName(child, localName, results);
        }
    }

    private static XmlNode FindFirstDescendantByLocalName(XmlNode node, string localName)
    {
        if (node == null)
        {
            return null;
        }
        if (node.LocalName == localName)
        {
            return node;
        }
        foreach (XmlNode child in node.ChildNodes)
        {
            XmlNode found = FindFirstDescendantByLocalName(child, localName);
            if (found != null)
            {
                return found;
            }
        }
        return null;
    }

    private static string GetAttributeValue(XmlNode node, string localName)
    {
        if (node == null || node.Attributes == null)
        {
            return "";
        }
        foreach (XmlAttribute attribute in node.Attributes)
        {
            if (attribute.LocalName == localName || attribute.Name == localName)
            {
                return attribute.Value;
            }
        }
        return "";
    }

    private static int GetColumnIndex(string cellReference)
    {
        if (String.IsNullOrEmpty(cellReference))
        {
            return 0;
        }

        int result = 0;
        foreach (char ch in cellReference.ToUpperInvariant())
        {
            if (ch < 'A' || ch > 'Z')
            {
                break;
            }
            if (result > (MaxExcelColumn - (ch - 'A' + 1)) / 26)
            {
                throw new InvalidDataException("Excel cell reference exceeds column XFD: " + cellReference);
            }
            result = (result * 26) + (ch - 'A' + 1);
        }
        return result;
    }

    private static string CollapseCellText(string text)
    {
        if (text == null)
        {
            return "";
        }
        return text.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ").Replace("\t", " ").Trim();
    }

    private static string NormalizeWorkbookTarget(string target)
    {
        if (String.IsNullOrEmpty(target))
        {
            return target;
        }
        target = target.Replace('\\', '/');
        if (target.StartsWith("/", StringComparison.Ordinal))
        {
            target = target.TrimStart('/');
        }
        else if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
        {
            target = "xl/" + target;
        }
        return NormalizePackagePath(target);
    }

    private static string NormalizePackagePath(string path)
    {
        return (path ?? "").Replace('\\', '/').TrimStart('/');
    }

    private static XmlDocument LoadEntryXml(ZipArchiveEntry entry)
    {
        if (entry == null)
        {
            throw new ArgumentNullException("entry");
        }
        if (entry.Length < 0 || entry.Length > MaxXmlEntryBytes)
        {
            throw new InvalidDataException("OpenXML part exceeds the 64 MiB XML limit: " + entry.FullName);
        }

        XmlDocument document = new XmlDocument();
        document.PreserveWhitespace = false;
        document.XmlResolver = null;
        using (Stream stream = entry.Open())
        using (XmlReader reader = XmlReader.Create(stream, CreateSecureXmlSettings()))
        {
            document.Load(reader);
        }
        return document;
    }

    private static XmlReaderSettings CreateSecureXmlSettings()
    {
        XmlReaderSettings settings = new XmlReaderSettings();
        settings.DtdProcessing = DtdProcessing.Prohibit;
        settings.XmlResolver = null;
        settings.MaxCharactersFromEntities = 0;
        settings.MaxCharactersInDocument = MaxXmlEntryBytes;
        settings.CloseInput = false;
        return settings;
    }

    private static void ValidateOpenXmlArchive(ZipArchive archive)
    {
        if (archive.Entries.Count > MaxOpenXmlEntries)
        {
            throw new InvalidDataException("OpenXML package exceeds the 10000-entry limit.");
        }

        long xmlBytes = 0;
        HashSet<string> names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            string name = NormalizePackagePath(entry.FullName);
            if (!names.Add(name))
            {
                throw new InvalidDataException("OpenXML package contains a duplicate part name: " + name);
            }
            if (!name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                !name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }
            if (entry.Length < 0 || entry.Length > MaxXmlEntryBytes)
            {
                throw new InvalidDataException("OpenXML part exceeds the 64 MiB XML limit: " + name);
            }
            if (xmlBytes > MaxOpenXmlBytes - entry.Length)
            {
                throw new InvalidDataException("OpenXML package exceeds the 256 MiB XML limit.");
            }
            xmlBytes += entry.Length;
        }
    }

    private static ZipArchiveEntry GetEntry(ZipArchive archive, string name)
    {
        string normalizedName = NormalizePackagePath(name);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            if (String.Equals(NormalizePackagePath(entry.FullName), normalizedName, StringComparison.OrdinalIgnoreCase))
            {
                return entry;
            }
        }
        return null;
    }

    private static List<ZipArchiveEntry> GetMatchingEntries(ZipArchive archive, string pattern)
    {
        Regex regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        List<ZipArchiveEntry> entries = new List<ZipArchiveEntry>();
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            string name = NormalizePackagePath(entry.FullName);
            if (regex.IsMatch(name))
            {
                entries.Add(entry);
            }
        }
        return entries;
    }

    private static int CompareEntryNumber(ZipArchiveEntry left, ZipArchiveEntry right)
    {
        int leftNumber = GetLastNumber(left.FullName);
        int rightNumber = GetLastNumber(right.FullName);
        int numberCompare = leftNumber.CompareTo(rightNumber);
        if (numberCompare != 0)
        {
            return numberCompare;
        }
        return StringComparer.OrdinalIgnoreCase.Compare(left.FullName, right.FullName);
    }

    private static int GetLastNumber(string text)
    {
        if (String.IsNullOrEmpty(text))
        {
            return 0;
        }
        MatchCollection matches = Regex.Matches(text, "[0-9]+");
        if (matches.Count == 0)
        {
            return 0;
        }
        int value;
        if (Int32.TryParse(matches[matches.Count - 1].Value, out value))
        {
            return value;
        }
        return 0;
    }

    private static void AppendLineIfNeeded(StringBuilder builder)
    {
        if (builder.Length == 0)
        {
            return;
        }
        string current = builder.ToString();
        if (!current.EndsWith("\n", StringComparison.Ordinal))
        {
            builder.AppendLine();
        }
    }

    private static string TrimEndLines(string text)
    {
        if (text == null)
        {
            return "";
        }
        return text.TrimEnd('\r', '\n', '\t', ' ');
    }
}
