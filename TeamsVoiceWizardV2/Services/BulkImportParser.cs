using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TeamsVoiceWizardV2.Models;

namespace TeamsVoiceWizardV2.Services;

/// <summary>
/// Parses bulk import CSV or XLSX files into a list of <see cref="BulkImportRow"/>.
/// Expected columns (case-insensitive, order-independent):
///   PhoneNumber | UserPrincipalName | DialPlan | VoiceRoutingPolicy
/// </summary>
public static class BulkImportParser
{
    private static readonly string[] PhoneNumberHeaders        = ["phonenumber", "phone", "phone number", "telephonenumber"];
    private static readonly string[] UpnHeaders                = ["userprincipalname", "upn", "email", "user"];
    private static readonly string[] DialPlanHeaders           = ["dialplan", "dial plan", "tenantdialplan"];
    private static readonly string[] VrpHeaders                = ["voiceroutingpolicy", "voice routing policy", "voicerouting", "vrp"];

    // ── CSV ───────────────────────────────────────────────────────────────────

    /// <summary>Parses a CSV stream. The first non-empty line is treated as a header row.</summary>
    public static List<BulkImportRow> ParseCsv(Stream stream)
    {
        var rows = new List<BulkImportRow>();
        using var reader = new StreamReader(stream, leaveOpen: true);

        string? headerLine = null;
        while (headerLine is null || headerLine.Trim().Length == 0)
        {
            headerLine = reader.ReadLine();
            if (headerLine is null) return rows; // empty file
        }

        var headers = ParseCsvLine(headerLine)
            .Select(h => h.Trim().ToLowerInvariant())
            .ToList();

        int phoneIdx = FindColumn(headers, PhoneNumberHeaders);
        int upnIdx   = FindColumn(headers, UpnHeaders);
        int dpIdx    = FindColumn(headers, DialPlanHeaders);
        int vrpIdx   = FindColumn(headers, VrpHeaders);

        if (phoneIdx < 0 || upnIdx < 0)
            throw new InvalidOperationException(
                "CSV must have 'PhoneNumber' and 'UserPrincipalName' columns.");

        int fileRow = 1;
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            fileRow++;
            if (string.IsNullOrWhiteSpace(line)) continue;

            var cells = ParseCsvLine(line);

            rows.Add(new BulkImportRow
            {
                RowNumber         = fileRow,
                PhoneNumber       = GetCell(cells, phoneIdx).Trim(),
                UserPrincipalName = GetCell(cells, upnIdx).Trim(),
                DialPlan          = dpIdx  >= 0 ? NullIfEmpty(GetCell(cells, dpIdx))  : null,
                VoiceRoutingPolicy= vrpIdx >= 0 ? NullIfEmpty(GetCell(cells, vrpIdx)) : null,
            });
        }

        return rows;
    }

    // ── XLSX ──────────────────────────────────────────────────────────────────

    /// <summary>Parses the first worksheet of an XLSX stream.</summary>
    public static List<BulkImportRow> ParseXlsx(Stream stream)
    {
        var rows = new List<BulkImportRow>();

        using var doc = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart  = doc.WorkbookPart
            ?? throw new InvalidOperationException("XLSX file has no workbook part.");

        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;

        var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault()
            ?? throw new InvalidOperationException("XLSX file has no worksheets.");

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()
            ?? throw new InvalidOperationException("Worksheet has no data.");

        var xlRows = sheetData.Elements<Row>().ToList();
        if (xlRows.Count == 0) return rows;

        // ── Header row ────────────────────────────────────────────────────────
        var headerRow = xlRows[0];
        var headerCells = GetRowValues(headerRow, sharedStrings);
        var headers = headerCells.Select(h => h.Trim().ToLowerInvariant()).ToList();

        int phoneIdx = FindColumn(headers, PhoneNumberHeaders);
        int upnIdx   = FindColumn(headers, UpnHeaders);
        int dpIdx    = FindColumn(headers, DialPlanHeaders);
        int vrpIdx   = FindColumn(headers, VrpHeaders);

        if (phoneIdx < 0 || upnIdx < 0)
            throw new InvalidOperationException(
                "XLSX must have 'PhoneNumber' and 'UserPrincipalName' columns.");

        // ── Data rows ─────────────────────────────────────────────────────────
        for (int i = 1; i < xlRows.Count; i++)
        {
            var cells = GetRowValues(xlRows[i], sharedStrings);
            if (cells.All(string.IsNullOrWhiteSpace)) continue;

            rows.Add(new BulkImportRow
            {
                RowNumber          = i + 1, // 1-based, matching Excel row numbers (header = row 1)
                PhoneNumber        = GetCell(cells, phoneIdx).Trim(),
                UserPrincipalName  = GetCell(cells, upnIdx).Trim(),
                DialPlan           = dpIdx  >= 0 ? NullIfEmpty(GetCell(cells, dpIdx))  : null,
                VoiceRoutingPolicy = vrpIdx >= 0 ? NullIfEmpty(GetCell(cells, vrpIdx)) : null,
            });
        }

        return rows;
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private static int FindColumn(List<string> headers, string[] candidates)
    {
        foreach (var candidate in candidates)
        {
            int idx = headers.IndexOf(candidate);
            if (idx >= 0) return idx;
        }
        return -1;
    }

    private static string GetCell(List<string> cells, int idx) =>
        idx >= 0 && idx < cells.Count ? cells[idx] : "";

    private static string? NullIfEmpty(string value) =>
        string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    /// <summary>Parses a single CSV line, respecting double-quoted fields.</summary>
    private static List<string> ParseCsvLine(string line)
    {
        var fields = new List<string>();
        var current = new System.Text.StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];

            if (c == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    current.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (c == ',' && !inQuotes)
            {
                fields.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(c);
            }
        }

        fields.Add(current.ToString());
        return fields;
    }

    /// <summary>Returns the string values of all cells in an XLSX row, in column order.</summary>
    private static List<string> GetRowValues(Row row, SharedStringTable? sharedStrings)
    {
        var cells = row.Elements<Cell>().ToList();
        if (cells.Count == 0) return [];

        // Cells may be sparse (A1, C1 with no B1) — expand to a dense list by column index
        int maxCol = cells.Select(c => ColumnIndex(c.CellReference?.Value)).Max();
        var values = new List<string>(new string[maxCol + 1]);

        foreach (var cell in cells)
        {
            int idx = ColumnIndex(cell.CellReference?.Value);
            if (idx < 0) continue;
            values[idx] = GetCellValue(cell, sharedStrings);
        }

        return values;
    }

    private static string GetCellValue(Cell cell, SharedStringTable? sharedStrings)
    {
        var raw = cell.CellValue?.Text ?? "";

        if (cell.DataType?.Value == CellValues.SharedString && sharedStrings is not null)
        {
            if (int.TryParse(raw, out int idx))
                return sharedStrings.ElementAt(idx).InnerText;
        }

        return raw;
    }

    /// <summary>Converts an Excel cell reference (e.g. "C5") to a zero-based column index.</summary>
    private static int ColumnIndex(string? cellRef)
    {
        if (string.IsNullOrWhiteSpace(cellRef)) return -1;

        int col = 0;
        foreach (char ch in cellRef)
        {
            if (!char.IsLetter(ch)) break;
            col = col * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
        }

        return col - 1;
    }
}
