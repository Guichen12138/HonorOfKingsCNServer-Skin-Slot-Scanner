using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace SkinViewer;

internal sealed record SkinRecord(
    int HeroId,
    string HeroName,
    int SkinId,
    string SkinName,
    string LastModified);

internal sealed class SheetSnapshot
{
    public int HeroCount { get; init; }
    public int SkinCount { get; init; }
    public Dictionary<(int HeroId, int SkinId), SkinRecord> Skins { get; init; } = new();
    public Dictionary<int, string> HeroNames { get; init; } = new();
    public SkinRecord? LatestSkin { get; init; }
}

internal sealed class SheetDiff
{
    public List<SkinRecord> Added { get; } = new();
    public List<(SkinRecord OldRec, SkinRecord NewRec)> Named { get; } = new();
    public List<(SkinRecord OldRec, SkinRecord NewRec)> Renamed { get; } = new();
    public List<(SkinRecord OldRec, SkinRecord NewRec)> TimeUpdated { get; } = new();
    public List<SkinRecord> Removed { get; } = new();
    public List<(int HeroId, string OldName, string NewName)> HeroNamed { get; } = new();

    public int TotalChangeCount =>
        Added.Count + Named.Count + Renamed.Count + TimeUpdated.Count + Removed.Count + HeroNamed.Count;

    public bool HasAnyChange => TotalChangeCount > 0;
}

internal static class XlsxParser
{
    public static bool IsValidXlsx(byte[] data)
    {
        if (data is null || data.Length < 1024) return false;
        if (data[0] != (byte)'P' || data[1] != (byte)'K') return false;

        try
        {
            using var z = new ZipArchive(new MemoryStream(data), ZipArchiveMode.Read);
            bool hasWorkbook = z.GetEntry("xl/workbook.xml") is not null
                || z.Entries.Any(e => string.Equals(e.FullName.TrimStart('/'), "xl/workbook.xml", StringComparison.OrdinalIgnoreCase));
            bool hasSheet = z.GetEntry("xl/worksheets/sheet1.xml") is not null
                || z.Entries.Any(e => e.FullName.TrimStart('/').StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                                   && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
            return hasWorkbook && hasSheet;
        }
        catch
        {
            return false;
        }
    }

    public static SheetSnapshot Parse(string path) => Parse(File.ReadAllBytes(path));

    public static SheetSnapshot Parse(byte[] xlsxBytes)
    {
        using var zip = new ZipArchive(new MemoryStream(xlsxBytes), ZipArchiveMode.Read);

        // 1) 读取可选的 sharedStrings.xml（Excel 保存时用；openpyxl 保存时常用 inlineStr）
        var sharedStrings = new List<string>();
        var ssEntry = zip.GetEntry("xl/sharedStrings.xml")
            ?? zip.Entries.FirstOrDefault(e => string.Equals(e.FullName.TrimStart('/'), "xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase));
        if (ssEntry is not null)
        {
            using var ssStream = ssEntry.Open();
            var ssDoc = XDocument.Load(ssStream);
            XNamespace ns = ssDoc.Root?.Name.Namespace ?? XNamespace.None;
            foreach (var si in ssDoc.Descendants(ns + "si"))
            {
                string val = string.Concat(si.Descendants(ns + "t").Select(t => t.Value));
                sharedStrings.Add(val);
            }
        }

        // 2) 读取第一个工作表 XML
        var sheetEntry = zip.GetEntry("xl/worksheets/sheet1.xml")
            ?? zip.Entries.FirstOrDefault(e => e.FullName.TrimStart('/').StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                                            && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidDataException("Excel 内未找到工作表 XML。");

        using var sheetStream = sheetEntry.Open();
        var doc = XDocument.Load(sheetStream);
        XNamespace mainNs = doc.Root?.Name.Namespace ?? XNamespace.None;

        var skinNameCols = new Dictionary<int, int>();
        var lmCols = new Dictionary<int, int>();
        var skins = new Dictionary<(int, int), SkinRecord>();
        var heroNames = new Dictionary<int, string>();
        int heroCount = 0;
        SkinRecord? latest = null;
        bool headerParsed = false;

        foreach (var rowEl in doc.Descendants(mainNs + "row"))
        {
            var rowCells = ReadRowCells(rowEl, mainNs, sharedStrings);
            if (rowCells.Count == 0) continue;

            if (!headerParsed)
            {
                headerParsed = true;
                foreach (var (colIdx, cell) in rowCells)
                {
                    string h = cell.Value.Trim();
                    if (h.EndsWith("-皮肤名", StringComparison.Ordinal))
                    {
                        if (int.TryParse(h.Split('-')[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int sid))
                            skinNameCols[sid] = colIdx;
                    }
                    else if (h.EndsWith("-Last-Modified", StringComparison.Ordinal))
                    {
                        if (int.TryParse(h.Split('-')[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int sid))
                            lmCols[sid] = colIdx;
                    }
                }
                continue;
            }

            if (!rowCells.TryGetValue(0, out var hidCell) ||
                !int.TryParse(hidCell.Value.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int heroId))
            {
                continue;
            }

            string heroName = rowCells.TryGetValue(1, out var hnCell) ? hnCell.Value.Trim() : "";

            // 与扫描器口径保持一致：00 原皮槽有 Last-Modified 才算已上线英雄
            if (lmCols.TryGetValue(0, out int firstLmCol))
            {
                if (!rowCells.TryGetValue(firstLmCol, out var firstLm) || string.IsNullOrWhiteSpace(firstLm.Value))
                {
                    continue;
                }
            }

            heroCount++;
            heroNames[heroId] = string.IsNullOrEmpty(heroName) ? $"{heroId:D3}" : heroName;

            foreach (var (sid, lmCol) in lmCols.OrderBy(kv => kv.Key))
            {
                if (!rowCells.TryGetValue(lmCol, out var lmCell) || string.IsNullOrWhiteSpace(lmCell.Value))
                    continue;

                string lm = NormalizeLastModified(lmCell);
                string sname = "";
                if (skinNameCols.TryGetValue(sid, out int snCol) && rowCells.TryGetValue(snCol, out var snCell))
                {
                    sname = snCell.Value.Trim();
                }

                var rec = new SkinRecord(heroId, heroNames[heroId], sid, sname, lm);
                skins[(heroId, sid)] = rec;

                if (latest is null || string.CompareOrdinal(rec.LastModified, latest.LastModified) > 0)
                {
                    latest = rec;
                }
            }
        }

        return new SheetSnapshot
        {
            HeroCount = heroCount,
            SkinCount = skins.Count,
            Skins = skins,
            HeroNames = heroNames,
            LatestSkin = latest,
        };
    }

    public static SheetDiff Compare(SheetSnapshot local, SheetSnapshot remote)
    {
        var diff = new SheetDiff();

        foreach (var (hid, rHeroName) in remote.HeroNames.OrderBy(kv => kv.Key))
        {
            if (local.HeroNames.TryGetValue(hid, out string? lHeroName)
                && !string.Equals(lHeroName, rHeroName, StringComparison.Ordinal))
            {
                diff.HeroNamed.Add((hid, lHeroName, rHeroName));
            }
        }

        var allKeys = local.Skins.Keys.Union(remote.Skins.Keys)
            .OrderBy(k => k.HeroId)
            .ThenBy(k => k.SkinId);

        foreach (var key in allKeys)
        {
            bool inLocal = local.Skins.TryGetValue(key, out var l);
            bool inRemote = remote.Skins.TryGetValue(key, out var r);

            if (!inLocal && inRemote && r is not null)
            {
                diff.Added.Add(r);
            }
            else if (inLocal && !inRemote && l is not null)
            {
                diff.Removed.Add(l);
            }
            else if (inLocal && inRemote && l is not null && r is not null)
            {
                if (string.IsNullOrEmpty(l.SkinName) && !string.IsNullOrEmpty(r.SkinName))
                {
                    diff.Named.Add((l, r));
                }
                else if (!string.Equals(l.SkinName, r.SkinName, StringComparison.Ordinal))
                {
                    diff.Renamed.Add((l, r));
                }

                if (!string.Equals(l.LastModified, r.LastModified, StringComparison.Ordinal))
                {
                    diff.TimeUpdated.Add((l, r));
                }
            }
        }

        return diff;
    }

    private static string NormalizeLastModified((string Value, bool IsNumericCell) cell)
    {
        string raw = cell.Value.Trim();
        if (cell.IsNumericCell
            && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double oaDate)
            && oaDate is > 35000.0 and < 80000.0)
        {
            try
            {
                return DateTime.FromOADate(oaDate).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }
            catch { }
        }
        return raw;
    }

    private static Dictionary<int, (string Value, bool IsNumericCell)> ReadRowCells(
        XElement rowEl,
        XNamespace ns,
        List<string> sharedStrings)
    {
        var cells = new Dictionary<int, (string Value, bool IsNumericCell)>();
        int nextCol = 0;

        foreach (var c in rowEl.Elements(ns + "c"))
        {
            string? cellRef = c.Attribute("r")?.Value;
            int colIdx = cellRef is not null ? ParseColumnIndex(cellRef) : nextCol;
            nextCol = colIdx + 1;

            string? cellType = c.Attribute("t")?.Value;
            string value = "";
            bool isNumeric = false;

            if (cellType == "inlineStr")
            {
                var isEl = c.Element(ns + "is");
                if (isEl is not null)
                {
                    value = string.Concat(isEl.Descendants(ns + "t").Select(t => t.Value));
                }
            }
            else if (cellType == "s")
            {
                string? rawIdx = c.Element(ns + "v")?.Value;
                if (int.TryParse(rawIdx, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idx)
                    && idx >= 0 && idx < sharedStrings.Count)
                {
                    value = sharedStrings[idx];
                }
            }
            else if (cellType == "str")
            {
                value = c.Element(ns + "v")?.Value ?? "";
            }
            else
            {
                value = c.Element(ns + "v")?.Value ?? "";
                isNumeric = string.IsNullOrEmpty(cellType) || cellType == "n";
            }

            if (!string.IsNullOrEmpty(value))
            {
                cells[colIdx] = (value, isNumeric);
            }
        }

        return cells;
    }

    private static int ParseColumnIndex(string cellRef)
    {
        int col = 0;
        foreach (char ch in cellRef)
        {
            if (ch >= 'A' && ch <= 'Z')
            {
                col = col * 26 + (ch - 'A' + 1);
            }
            else if (ch >= 'a' && ch <= 'z')
            {
                col = col * 26 + (ch - 'a' + 1);
            }
            else
            {
                break;
            }
        }
        return Math.Max(0, col - 1);
    }
}
