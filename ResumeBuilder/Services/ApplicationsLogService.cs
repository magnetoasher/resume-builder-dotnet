using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using ClosedXML.Excel;
using ResumeBuilder.Models;

namespace ResumeBuilder.Services;

public class ApplicationsLogService
{
    private static readonly string[] RequiredHeaders =
    [
        "Timestamp",
        "Company",
        "Role",
        "Job URL",
        "Job Description",
        "Resume Path",
        "Profile"
    ];

    public void AppendEntry(string path, ApplicationLogEntry entry)
    {
        const int maxAttempts = 3;
        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                AppendEntryInternal(path, entry);
                return;
            }
            catch (IOException) when (attempt < maxAttempts)
            {
                Thread.Sleep(250);
            }
        }

        // Final attempt throws the original exception.
        AppendEntryInternal(path, entry);
    }

    private static void AppendEntryInternal(string path, ApplicationLogEntry entry)
    {
        var exists = File.Exists(path);
        using var workbook = exists ? new XLWorkbook(path) : new XLWorkbook();
        var worksheet = workbook.Worksheets.FirstOrDefault() ?? workbook.AddWorksheet("Applications");

        var headerMap = EnsureHeaders(worksheet);
        var nextRow = (worksheet.LastRowUsed()?.RowNumber() ?? 1) + 1;

        worksheet.Cell(nextRow, headerMap["Timestamp"]).Value = entry.Timestamp;
        worksheet.Cell(nextRow, headerMap["Company"]).Value = entry.Company;
        worksheet.Cell(nextRow, headerMap["Role"]).Value = entry.Role;
        worksheet.Cell(nextRow, headerMap["Job URL"]).Value = entry.JobUrl;
        worksheet.Cell(nextRow, headerMap["Job Description"]).Value = entry.JobDescription;
        worksheet.Cell(nextRow, headerMap["Resume Path"]).Value = entry.ResumePath;
        worksheet.Cell(nextRow, headerMap["Profile"]).Value = entry.ProfileName;

        workbook.SaveAs(path);
    }

    private static Dictionary<string, int> EnsureHeaders(IXLWorksheet worksheet)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var lastCol = worksheet.Row(1).LastCellUsed()?.Address.ColumnNumber ?? 0;

        for (var col = 1; col <= lastCol; col++)
        {
            var header = worksheet.Cell(1, col).GetString().Trim();
            if (string.IsNullOrWhiteSpace(header) || map.ContainsKey(header))
            {
                continue;
            }

            map[header] = col;
        }

        var nextColumn = lastCol + 1;
        foreach (var header in RequiredHeaders)
        {
            if (map.ContainsKey(header))
            {
                continue;
            }

            worksheet.Cell(1, nextColumn).Value = header;
            map[header] = nextColumn;
            nextColumn++;
        }

        return map;
    }
}
