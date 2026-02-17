using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ResumeBuilder.Models;

namespace ResumeBuilder.Services;

public class ApplicationsQueryService
{
    private static readonly string[] Headers =
    [
        "Timestamp",
        "Company",
        "Role",
        "Job URL",
        "Job Description",
        "Resume Path",
        "Profile"
    ];

    public IReadOnlyList<ApplicationRecord> Load(string path)
    {
        var records = new List<ApplicationRecord>();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return records;
        }

        using var workbook = new XLWorkbook(path);
        var worksheet = workbook.Worksheets.FirstOrDefault();
        if (worksheet == null)
        {
            return records;
        }

        var headerMap = BuildHeaderMap(worksheet);
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
        for (var row = 2; row <= lastRow; row++)
        {
            var timestamp = GetCell(worksheet, headerMap, "Timestamp", row);
            var company = GetCell(worksheet, headerMap, "Company", row);
            var role = GetCell(worksheet, headerMap, "Role", row);
            var jobUrl = GetCell(worksheet, headerMap, "Job URL", row);
            var jobDescription = GetCell(worksheet, headerMap, "Job Description", row);
            var resumePath = GetCell(worksheet, headerMap, "Resume Path", row);
            var profileName = GetCell(worksheet, headerMap, "Profile", row);

            if (string.IsNullOrWhiteSpace(timestamp) &&
                string.IsNullOrWhiteSpace(company) &&
                string.IsNullOrWhiteSpace(role) &&
                string.IsNullOrWhiteSpace(jobUrl) &&
                string.IsNullOrWhiteSpace(jobDescription) &&
                string.IsNullOrWhiteSpace(resumePath) &&
                string.IsNullOrWhiteSpace(profileName))
            {
                continue;
            }

            DateTime? timestampValue = null;
            if (DateTime.TryParse(timestamp, out var parsedTimestamp))
            {
                timestampValue = parsedTimestamp;
            }

            records.Add(new ApplicationRecord
            {
                Timestamp = timestamp,
                TimestampValue = timestampValue,
                Company = company,
                Role = role,
                JobUrl = jobUrl,
                JobDescription = jobDescription,
                ResumePath = resumePath,
                ProfileName = profileName
            });
        }

        return records
            .OrderByDescending(record => record.TimestampValue ?? DateTime.MinValue)
            .ThenByDescending(record => record.Timestamp)
            .ToList();
    }

    private static Dictionary<string, int> BuildHeaderMap(IXLWorksheet worksheet)
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

        // Support files that miss newer headers by still returning empty values for them.
        foreach (var header in Headers)
        {
            _ = map.ContainsKey(header);
        }

        return map;
    }

    private static string GetCell(IXLWorksheet worksheet, IReadOnlyDictionary<string, int> headerMap, string header, int row)
    {
        if (!headerMap.TryGetValue(header, out var col))
        {
            return string.Empty;
        }

        return worksheet.Cell(row, col).GetString().Trim();
    }
}
