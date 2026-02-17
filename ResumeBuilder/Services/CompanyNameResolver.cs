using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

namespace ResumeBuilder.Services;

public readonly record struct JobPostDetails(string Company, string Role);

public static class CompanyNameResolver
{
    private const string RoleKeywordPattern = "engineer|developer|architect|manager|manger|analyst|scientist|consultant|specialist|administrator|designer|lead|director|officer";

    private static readonly HashSet<string> IgnoredDescriptionLines = new(StringComparer.OrdinalIgnoreCase)
    {
        "apply",
        "job details",
        "skills",
        "summary",
        "profile insights",
        "job type",
        "benefits",
        "full job description",
        "responsibilities",
        "skillset",
        "education & requirements",
        "education and requirements",
        "fitment",
        "remote",
        "full-time",
        "contract w2",
        "no travel required",
        "depends on experience"
    };

    private static readonly HashSet<string> RoleKeywords = new(StringComparer.OrdinalIgnoreCase)
    {
        "engineer",
        "developer",
        "architect",
        "manager",
        "manger",
        "analyst",
        "scientist",
        "consultant",
        "specialist",
        "administrator",
        "designer",
        "lead",
        "director",
        "officer"
    };

    private static readonly HashSet<string> GenericJobHosts = new(StringComparer.OrdinalIgnoreCase)
    {
        "dice.com",
        "indeed.com",
        "linkedin.com",
        "ziprecruiter.com",
        "glassdoor.com",
        "monster.com",
        "careerbuilder.com",
        "simplyhired.com",
        "talent.com",
        "wellfound.com",
        "jobcase.com"
    };

    public static JobPostDetails InferJobPostDetails(string jobUrl, string jobDescription)
    {
        var lines = ExtractLines(jobDescription);
        var role = InferRoleFromDescription(lines, jobDescription);
        var company = InferCompanyFromDescription(lines, jobDescription, role);

        if (string.IsNullOrWhiteSpace(company))
        {
            company = InferFromUrl(jobUrl);
        }

        return new JobPostDetails(company, role);
    }

    public static string InferCompanyName(string jobUrl, string jobDescription)
    {
        return InferJobPostDetails(jobUrl, jobDescription).Company;
    }

    public static string InferRoleName(string jobUrl, string jobDescription)
    {
        return InferJobPostDetails(jobUrl, jobDescription).Role;
    }

    private static string InferRoleFromDescription(IReadOnlyList<string> lines, string jobDescription)
    {
        if (string.IsNullOrWhiteSpace(jobDescription))
        {
            return string.Empty;
        }

        var topLines = lines.Take(32).ToList();

        // Prefer headline-style role lines near the top of the pasted JD.
        foreach (var line in topLines)
        {
            var cleaned = CleanRoleCandidate(line);
            if (IsLikelyRole(cleaned))
            {
                return cleaned;
            }
        }

        var explicitRole = Regex.Match(jobDescription, @"(?im)^\s*(?:role|position|title)\s*:\s*(.+)$");
        if (explicitRole.Success)
        {
            var cleaned = CleanRoleCandidate(explicitRole.Groups[1].Value);
            if (IsLikelyRole(cleaned))
            {
                return cleaned;
            }
        }

        foreach (var line in lines.Take(60))
        {
            var cleaned = CleanRoleCandidate(line);
            if (!IsLikelyRole(cleaned))
            {
                continue;
            }

            if (LooksLikeCompany(cleaned, role: string.Empty, duplicateCount: 1))
            {
                continue;
            }

            return cleaned;
        }

        return string.Empty;
    }

    private static string InferCompanyFromDescription(IReadOnlyList<string> lines, string jobDescription, string role)
    {
        if (string.IsNullOrWhiteSpace(jobDescription))
        {
            return string.Empty;
        }

        var explicitCompany = Regex.Match(jobDescription, @"(?im)^\s*company\s*[:\-]\s*(.+)$");
        if (explicitCompany.Success)
        {
            return CleanCompanyName(explicitCompany.Groups[1].Value);
        }

        var joinAt = Regex.Match(jobDescription, @"(?is)\bjoin(?:\s+our\s+[A-Za-z0-9&'\- ]+)?\s+(?:team\s+)?at\s+([A-Z][A-Za-z0-9&.,'\- ]{1,80})(?:[\r\n.,]|$)");
        if (joinAt.Success)
        {
            return CleanCompanyName(joinAt.Groups[1].Value);
        }

        var topLines = lines.Take(40).ToList();
        var duplicateCounts = topLines
            .Select(CleanCompanyName)
            .Where(candidate => !string.IsNullOrWhiteSpace(candidate))
            .GroupBy(candidate => candidate, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

        // If we found role from headline, company is usually right above/below it.
        var roleIndex = FindRoleLineIndex(topLines, role);
        if (roleIndex >= 0)
        {
            for (var offset = 1; offset <= 6; offset++)
            {
                var nextIndex = roleIndex + offset;
                if (nextIndex < topLines.Count)
                {
                    var nextLine = topLines[nextIndex];
                    var cleaned = CleanCompanyName(nextLine);
                    var duplicateCount = duplicateCounts.TryGetValue(cleaned, out var nextCount) ? nextCount : 1;
                    if (LooksLikeCompany(nextLine, role, duplicateCount))
                    {
                        return cleaned;
                    }
                }
            }

            for (var offset = 1; offset <= 6; offset++)
            {
                var prevIndex = roleIndex - offset;
                if (prevIndex >= 0)
                {
                    var prevLine = topLines[prevIndex];
                    var cleaned = CleanCompanyName(prevLine);
                    var duplicateCount = duplicateCounts.TryGetValue(cleaned, out var prevCount) ? prevCount : 1;
                    if (LooksLikeCompany(prevLine, role, duplicateCount))
                    {
                        return cleaned;
                    }
                }
            }
        }

        // Job boards often repeat company name lines (e.g., iTek People, Inc. twice).
        foreach (var line in topLines)
        {
            var cleaned = CleanCompanyName(line);
            var duplicateCount = duplicateCounts.TryGetValue(cleaned, out var count) ? count : 1;
            if (duplicateCount >= 2 && LooksLikeCompany(line, role, duplicateCount))
            {
                return cleaned;
            }
        }

        foreach (var line in topLines)
        {
            var cleaned = CleanCompanyName(line);
            var duplicateCount = duplicateCounts.TryGetValue(cleaned, out var count) ? count : 1;
            if (LooksLikeCompany(line, role, duplicateCount))
            {
                return cleaned;
            }
        }

        return string.Empty;
    }

    private static List<string> ExtractLines(string jobDescription)
    {
        if (string.IsNullOrWhiteSpace(jobDescription))
        {
            return [];
        }

        return jobDescription
            .Replace("\r\n", "\n")
            .Split('\n')
            .Select(WebUtility.HtmlDecode)
            .Select(line => Regex.Replace(line ?? string.Empty, @"\s+", " ").Trim())
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Where(line => line.Length <= 140)
            .ToList();
    }

    private static string InferFromUrl(string jobUrl)
    {
        if (string.IsNullOrWhiteSpace(jobUrl))
        {
            return string.Empty;
        }

        if (!Uri.TryCreate(jobUrl.Trim(), UriKind.Absolute, out var uri))
        {
            return string.Empty;
        }

        var host = uri.Host.ToLowerInvariant().Replace("www.", string.Empty);

        if (host.Contains("greenhouse.io", StringComparison.OrdinalIgnoreCase))
        {
            var segment = uri.AbsolutePath.Trim('/').Split('/').FirstOrDefault();
            return FormatName(segment);
        }

        if (host.Contains("lever.co", StringComparison.OrdinalIgnoreCase))
        {
            var segment = uri.AbsolutePath.Trim('/').Split('/').FirstOrDefault();
            return FormatName(segment);
        }

        if (host.Contains("myworkdayjobs", StringComparison.OrdinalIgnoreCase))
        {
            var segment = uri.AbsolutePath.Trim('/').Split('/').FirstOrDefault();
            return FormatName(segment);
        }

        if (GenericJobHosts.Any(generic => host.EndsWith(generic, StringComparison.OrdinalIgnoreCase)))
        {
            return string.Empty;
        }

        var parts = host.Split('.').Where(p => !string.IsNullOrWhiteSpace(p)).ToArray();
        if (parts.Length < 2)
        {
            return string.Empty;
        }

        var core = parts[^2];
        if (core is "jobs" or "careers" or "apply")
        {
            core = parts.Length >= 3 ? parts[^3] : core;
        }

        if (core is "dice" or "indeed" or "linkedin")
        {
            return string.Empty;
        }

        return FormatName(core);
    }

    private static bool IsLikelyRole(string value)
    {
        var candidate = CleanRoleCandidate(value);
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return false;
        }

        if (candidate.Length < 6 || candidate.Length > 120)
        {
            return false;
        }

        if (IgnoredDescriptionLines.Contains(candidate.Trim()))
        {
            return false;
        }

        if (ContainsLinkOrEmail(candidate))
        {
            return false;
        }

        if (Regex.IsMatch(candidate, @"(?i)\b(inc\.?|llc|ltd\.?|corp\.?)\b"))
        {
            return false;
        }

        if (Regex.IsMatch(candidate, @"(?i)^(posted|updated|remote|full[- ]?time|part[- ]?time|contract|benefits?)\b"))
        {
            return false;
        }

        if (Regex.IsMatch(candidate, @"(?i)\b(job\s*details|profile insights|full job description|qualifications)\b"))
        {
            return false;
        }

        return RoleKeywords.Any(keyword => candidate.Contains(keyword, StringComparison.OrdinalIgnoreCase));
    }

    private static bool LooksLikeCompany(string value, string role, int duplicateCount)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var line = value.Trim();
        if (line.Length < 2 || line.Length > 100)
        {
            return false;
        }

        if (IgnoredDescriptionLines.Contains(line))
        {
            return false;
        }

        if (ContainsLinkOrEmail(line))
        {
            return false;
        }

        if (string.Equals(line, "Confidential", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var normalizedRole = NormalizeForCompare(CleanRoleCandidate(role));
        var normalizedLineAsRole = NormalizeForCompare(CleanRoleCandidate(line));
        if (!string.IsNullOrWhiteSpace(normalizedRole) &&
            !string.IsNullOrWhiteSpace(normalizedLineAsRole) &&
            (string.Equals(normalizedLineAsRole, normalizedRole, StringComparison.Ordinal) ||
             normalizedLineAsRole.Contains(normalizedRole, StringComparison.Ordinal) ||
             normalizedRole.Contains(normalizedLineAsRole, StringComparison.Ordinal)))
        {
            return false;
        }

        if (Regex.IsMatch(line, @"(?i)^(role|job\s*details|summary|responsibilities|benefits|skillset|jd)\b"))
        {
            return false;
        }

        if (Regex.IsMatch(line, @"(?i)^\d+(?:\.\d+)?(?:\s*out\s+of\s+\d+(?:\.\d+)?)?(?:\s*stars?)?$"))
        {
            return false;
        }

        if (Regex.IsMatch(line, @"(?i)\bstars?\b|\bposted\b|\bupdated\b"))
        {
            return false;
        }

        if (Regex.IsMatch(line, @"(?i)\b(do you have experience|here(?:'|’)?s how|work location|work from home)\b"))
        {
            return false;
        }

        if (Regex.IsMatch(line, @"(?i)^\+?\s*show more$"))
        {
            return false;
        }

        if (Regex.IsMatch(line, @"(?i)^\(?\s*(required|preferred|highly desired)\s*\)?$"))
        {
            return false;
        }

        if (Regex.IsMatch(line, @"(?i)^(core backend|security\s*&\s*authentication|healthcare standards|integration\s*&\s*processing|infrastructure|development tools)\b"))
        {
            return false;
        }

        if (Regex.IsMatch(line, @"(?i)\$\s*\d"))
        {
            return false;
        }

        if (Regex.IsMatch(line, @"(?i)\b(full[- ]?time|part[- ]?time|contract|temporary|remote|hybrid|onsite)\b"))
        {
            return false;
        }

        if (Regex.IsMatch(line, @"(?i)\b\d+\+?\s*years?\b"))
        {
            return false;
        }

        if (IsLikelyRole(line))
        {
            return false;
        }

        var wordCount = line.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length;
        if (wordCount > 8)
        {
            return false;
        }

        if (duplicateCount >= 2)
        {
            return true;
        }

        if (Regex.IsMatch(line, @"(?i)\b(inc\.?|llc|ltd\.?|corp\.?|company|technologies|systems|labs|solutions|group)\b"))
        {
            return true;
        }

        if (line.Contains('.', StringComparison.Ordinal) || line.Contains(',', StringComparison.Ordinal))
        {
            return true;
        }

        return Regex.IsMatch(line, @"^[A-Za-z][A-Za-z0-9&'\- ]{1,60}$");
    }

    private static int FindRoleLineIndex(IReadOnlyList<string> lines, string role)
    {
        if (string.IsNullOrWhiteSpace(role))
        {
            return -1;
        }

        var target = NormalizeForCompare(CleanRoleCandidate(role));
        if (string.IsNullOrWhiteSpace(target))
        {
            return -1;
        }

        for (var i = 0; i < lines.Count; i++)
        {
            var current = NormalizeForCompare(CleanRoleCandidate(lines[i]));
            if (string.IsNullOrWhiteSpace(current))
            {
                continue;
            }

            if (string.Equals(current, target, StringComparison.Ordinal))
            {
                return i;
            }

            if (current.Contains(target, StringComparison.Ordinal) || target.Contains(current, StringComparison.Ordinal))
            {
                return i;
            }
        }

        return -1;
    }

    private static string NormalizeForCompare(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = Regex.Replace(value.ToLowerInvariant(), @"[^a-z0-9]+", " ").Trim();
        return Regex.Replace(normalized, @"\s+", " ");
    }

    private static bool ContainsLinkOrEmail(string value)
    {
        return value.Contains("http", StringComparison.OrdinalIgnoreCase)
            || value.Contains("www.", StringComparison.OrdinalIgnoreCase)
            || Regex.IsMatch(value, @"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", RegexOptions.IgnoreCase);
    }

    private static string CleanRoleCandidate(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return string.Empty;
        }

        var value = WebUtility.HtmlDecode(raw).Trim();
        value = Regex.Replace(value, @"(?i)^(?:role|position|title)\s*:\s*", string.Empty);
        value = Regex.Replace(value, @"(?i)\s*-\s*job\s+post.*$", string.Empty);
        value = Regex.Replace(value, @"(?i)\s*@\s*(remote|hybrid|onsite).*$", string.Empty);
        value = Regex.Replace(value, @"(?i)\s+\((?:w2|c2c|1099|contract)\b[^)]*\)", string.Empty);
        value = Regex.Replace(value, @"(?i)\s*[|•]\s*(posted|updated).*$", string.Empty);
        value = Regex.Replace(value, @"(?i)\s*[-|]\s*(remote|hybrid|onsite)\b.*$", string.Empty);
        value = Regex.Replace(value, @"(?i)\s+\((?:w2\s*only|c2c\s*only|contract\s*only)\)", string.Empty);

        var roleWithDecorators = Regex.Match(value, $@"(?i)^(?<core>.+?\b(?:{RoleKeywordPattern})\b)\s+with\b.+$");
        if (roleWithDecorators.Success)
        {
            value = roleWithDecorators.Groups["core"].Value;
        }

        value = Regex.Replace(value, @"\s+", " ").Trim();
        value = value.Trim('-', '|', ':', '.', ' ');
        return value;
    }

    private static string CleanCompanyName(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return string.Empty;
        }

        var value = WebUtility.HtmlDecode(raw).Trim();
        value = Regex.Replace(value, @"^(?i)company\s*[:\-]\s*", string.Empty);
        value = Regex.Replace(value, @"\s+", " ");
        value = value.Trim(' ', '-', ':');
        return value;
    }

    private static string FormatName(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return string.Empty;
        }

        var cleaned = raw.Replace('-', ' ').Replace('_', ' ').Trim();
        cleaned = Regex.Replace(cleaned, @"\s+", " ");
        if (cleaned.Length == 0)
        {
            return string.Empty;
        }

        return char.ToUpper(cleaned[0]) + cleaned[1..];
    }
}

