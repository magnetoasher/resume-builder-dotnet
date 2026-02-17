using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ResumeBuilder.Models;

namespace ResumeBuilder.Services;

public class ResumeGenerationService
{
    public async Task<ResumeContent> GenerateResumeAsync(
        Profile profile,
        string jobDescription,
        string candidateCorpus,
        string apiKey,
        AppSettings settings)
    {
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            throw new InvalidOperationException("OpenAI API key is missing.");
        }

        var openAi = new OpenAiService(apiKey);
        var years = ComputeYearsOfExperience(profile);
        var baseRoles = profile.Experience.Select(exp => new
        {
            company = exp.Company ?? string.Empty,
            title = exp.Title ?? string.Empty,
            dates = exp.Dates ?? string.Empty
        }).Cast<object>().ToList();

        var systemMessage = "You output strict JSON only. You are writing a professional resume. Never include phone numbers, emails, personal addresses, or new company names.";
        var prompt = BuildPrompt(jobDescription, candidateCorpus, years, baseRoles);

        var response = await openAi.CreateChatCompletionAsync(systemMessage, prompt, settings.Model, settings.Temperature, settings.TopP);
        var jsonText = OpenAiService.ExtractJson(response);

        var (payload, issues) = ParseAndValidate(jsonText, profile, years);
        if (issues.Count > 0)
        {
            var repairPrompt = BuildRepairPrompt(prompt, jsonText, issues);
            var repairResponse = await openAi.CreateChatCompletionAsync(systemMessage, repairPrompt, settings.Model, settings.Temperature, settings.TopP);
            var repairJson = OpenAiService.ExtractJson(repairResponse);
            (payload, issues) = ParseAndValidate(repairJson, profile, years);
            if (issues.Count > 0)
            {
                throw new Exception("AI response failed validation: " + string.Join("; ", issues));
            }
        }

        ApplySkillRealismRules(payload, jobDescription);
        ApplyBulletKeywordEmphasis(payload);
        return payload;
    }

    private static string BuildPrompt(string jobDescription, string candidateCorpus, string years, List<object> baseRoles)
    {
        var jd = (jobDescription ?? string.Empty).Trim();
        if (jd.Length > 7000)
        {
            jd = jd[..7000];
        }

        var corpus = (candidateCorpus ?? string.Empty).Trim();
        if (corpus.Length > 5000)
        {
            corpus = corpus[..5000];
        }

        return @"Mode: GENERATE NEW resume content.

Generate-mode rules:
- Rewrite summary, skills, and bullets for JD alignment while staying realistic and grounded in the provided base roles.
- Do not copy or mirror JD sentences verbatim or near-verbatim.
- Do not lift long JD phrases directly; paraphrase them into resume language.
- Avoid generic AI-sounding phrasing, buzzword stuffing, and empty claims.
- Use natural, concrete wording that sounds like a real candidate's resume.
- Candidate resume text is grounding context, not fixed output.
"
        + $@"

Output strict JSON only with this schema:
{{
  ""summary"": string,
  ""skills"": [""Category: item1, item2, ...""],
  ""experience"": [
    {{""company"": string, ""title"": string, ""dates"": string, ""bullets"": [string, ...]}}
  ]
}}

Rules (non-negotiable):
- Summary must be 4-5 sentences and start with '{years} years of experience'.
- Do NOT include phone numbers, emails, addresses, or links.
- Do NOT introduce company names that are not in base roles.
- Experience roles must match the base roles exactly in company/title/dates and count.
- Each role must have 5-6 bullets.
- Bullets must be JD-aligned and avoid wild claims. Prefer grounded wording.
- Bullets must not read like pasted JD responsibilities; convert them into candidate-focused accomplishments/responsibilities.
- In each experience bullet, bold 1-2 concrete technologies/tools using **double-asterisk** markers.
- Skills must have 3-6 category lines and each line must list 5-8 comma-separated items.
- Skills must include only technologies required/preferred by the JD (plus minimal direct companions).
- Avoid ""kitchen sink"" stacks. Do not list mutually parallel alternatives unless the JD asks for them.
- For frontend frameworks, include at most two among React/Angular/Vue.

Base roles (must match exactly):
{JsonSerializer.Serialize(baseRoles)}

Candidate resume text (for grounding if relevant):
{corpus}

Job description:
{jd}
";
    }

    private static string BuildRepairPrompt(string originalPrompt, string previousJson, List<string> issues)
    {
        var issueText = string.Join("\n- ", issues);
        return $@"{originalPrompt}

Your previous JSON output had problems:
- {issueText}

Previous output JSON:
{previousJson}

Return corrected JSON only.";
    }

    private static (ResumeContent, List<string>) ParseAndValidate(string jsonText, Profile profile, string years)
    {
        var issues = new List<string>();
        ResumeContent content = new();

        if (string.IsNullOrWhiteSpace(jsonText))
        {
            issues.Add("Empty JSON output.");
            return (content, issues);
        }

        JsonDocument doc;
        try
        {
            doc = JsonDocument.Parse(jsonText);
        }
        catch
        {
            issues.Add("Invalid JSON.");
            return (content, issues);
        }

        using (doc)
        {
            var root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                issues.Add("Top-level JSON must be an object.");
                return (content, issues);
            }

            var summary = root.TryGetProperty("summary", out var summaryEl) ? summaryEl.GetString() ?? string.Empty : string.Empty;
            if (string.IsNullOrWhiteSpace(summary))
            {
                issues.Add("Summary is missing.");
            }
            else
            {
                if (!summary.Trim().StartsWith($"{years} years of experience", StringComparison.OrdinalIgnoreCase))
                {
                    issues.Add("Summary must start with the required years-of-experience phrase.");
                }
                if (ContainsPersonalData(summary))
                {
                    issues.Add("Summary must not contain contact details or links.");
                }
                var sentences = Regex.Split(summary.Trim(), "(?<=[.!?])\\s+").Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                if (sentences.Count is < 4 or > 5)
                {
                    issues.Add("Summary must be 4-5 sentences.");
                }
            }

            var skills = new List<string>();
            if (root.TryGetProperty("skills", out var skillsEl) && skillsEl.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in skillsEl.EnumerateArray())
                {
                    if (item.ValueKind == JsonValueKind.String)
                    {
                        skills.Add(item.GetString() ?? string.Empty);
                    }
                }
            }
            if (skills.Count < 3)
            {
                issues.Add("Skills must include at least 3 category lines.");
            }
            else
            {
                foreach (var line in skills)
                {
                    if (!line.Contains(':'))
                    {
                        issues.Add("Each skills line must use 'Category: item1, item2' format.");
                        break;
                    }

                    var parts = line.Split(':', 2);
                    if (parts.Length != 2)
                    {
                        issues.Add("Each skills line must use 'Category: item1, item2' format.");
                        break;
                    }

                    var items = parts[1]
                        .Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                        .ToList();
                    if (items.Count is < 5 or > 8)
                    {
                        issues.Add("Each skills line must list 5-8 comma-separated items.");
                        break;
                    }
                }
            }

            var experience = new List<ResumeExperience>();
            if (root.TryGetProperty("experience", out var expEl) && expEl.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in expEl.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.Object)
                    {
                        continue;
                    }

                    var company = item.TryGetProperty("company", out var c) ? c.GetString() ?? string.Empty : string.Empty;
                    var title = item.TryGetProperty("title", out var t) ? t.GetString() ?? string.Empty : string.Empty;
                    var dates = item.TryGetProperty("dates", out var d) ? d.GetString() ?? string.Empty : string.Empty;

                    var bullets = new List<string>();
                    if (item.TryGetProperty("bullets", out var b) && b.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var bullet in b.EnumerateArray())
                        {
                            if (bullet.ValueKind == JsonValueKind.String)
                            {
                                bullets.Add(bullet.GetString() ?? string.Empty);
                            }
                        }
                    }

                    experience.Add(new ResumeExperience
                    {
                        Company = company,
                        Title = title,
                        Dates = dates,
                        Bullets = bullets
                    });
                }
            }

            if (experience.Count != profile.Experience.Count)
            {
                issues.Add("Experience roles must match the base profile count.");
            }
            else
            {
                for (var i = 0; i < profile.Experience.Count; i++)
                {
                    var baseRole = profile.Experience[i];
                    var role = experience[i];
                    if (!string.Equals(baseRole.Company?.Trim(), role.Company?.Trim(), StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(baseRole.Title?.Trim(), role.Title?.Trim(), StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(baseRole.Dates?.Trim(), role.Dates?.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        issues.Add("Experience roles must match the base profile company/title/dates exactly.");
                        break;
                    }

                    if (role.Bullets.Count is < 5 or > 6)
                    {
                        issues.Add("Each role must have 5-6 bullets.");
                        break;
                    }

                    if (role.Bullets.Any(ContainsPersonalData))
                    {
                        issues.Add("Bullets must not contain contact details or links.");
                        break;
                    }
                }
            }

            content = new ResumeContent
            {
                Summary = summary,
                Skills = skills,
                Experience = experience,
                Education = profile.Education ?? new List<Education>()
            };
        }

        return (content, issues);
    }

    private static void ApplySkillRealismRules(ResumeContent content, string jobDescription)
    {
        if (content.Skills.Count == 0)
        {
            return;
        }

        var jd = (jobDescription ?? string.Empty).ToLowerInvariant();
        var normalizedLines = new List<string>();

        foreach (var line in content.Skills)
        {
            if (string.IsNullOrWhiteSpace(line) || !line.Contains(':'))
            {
                continue;
            }

            var parts = line.Split(':', 2);
            var category = parts[0].Trim();
            var items = parts[1]
                .Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (items.Count == 0)
            {
                continue;
            }

            if (IsFrontendCategory(category))
            {
                items = PruneFrontendFrameworks(items, jd);
            }

            // Keep the list focused and readable.
            if (items.Count > 8)
            {
                items = items.Take(8).ToList();
            }

            normalizedLines.Add($"{category}: {string.Join(", ", items)}");
        }

        if (normalizedLines.Count > 0)
        {
            content.Skills = normalizedLines;
        }
    }

    private static bool IsFrontendCategory(string category)
    {
        var value = category?.Trim().ToLowerInvariant() ?? string.Empty;
        return value.Contains("front");
    }

    private static List<string> PruneFrontendFrameworks(List<string> items, string jdLower)
    {
        static string? MapFramework(string item)
        {
            var value = item.ToLowerInvariant();
            if (value.Contains("react"))
            {
                return "react";
            }

            if (value.Contains("angular"))
            {
                return "angular";
            }

            if (value.Contains("vue"))
            {
                return "vue";
            }

            return null;
        }

        var keysInOrder = items
            .Select(MapFramework)
            .Where(k => k != null)
            .Cast<string>()
            .Distinct()
            .ToList();

        if (keysInOrder.Count <= 2)
        {
            return items;
        }

        var keep = new List<string>();
        foreach (var key in keysInOrder)
        {
            if (keep.Count >= 2)
            {
                break;
            }

            if (jdLower.Contains(key, StringComparison.OrdinalIgnoreCase))
            {
                keep.Add(key);
            }
        }

        foreach (var key in keysInOrder)
        {
            if (keep.Count >= 2)
            {
                break;
            }

            if (!keep.Contains(key, StringComparer.OrdinalIgnoreCase))
            {
                keep.Add(key);
            }
        }

        return items
            .Where(item =>
            {
                var key = MapFramework(item);
                return key == null || keep.Contains(key, StringComparer.OrdinalIgnoreCase);
            })
            .ToList();
    }

    private static void ApplyBulletKeywordEmphasis(ResumeContent content)
    {
        var skillTerms = ExtractSkillTerms(content.Skills);
        if (skillTerms.Count == 0)
        {
            return;
        }

        foreach (var role in content.Experience)
        {
            for (var i = 0; i < role.Bullets.Count; i++)
            {
                var bullet = role.Bullets[i] ?? string.Empty;
                if (ContainsBoldToken(bullet))
                {
                    continue;
                }

                foreach (var term in skillTerms)
                {
                    if (TryBoldFirstOccurrence(bullet, term, out var updated))
                    {
                        bullet = updated;
                        break;
                    }
                }

                role.Bullets[i] = bullet;
            }
        }
    }

    private static List<string> ExtractSkillTerms(List<string> skills)
    {
        var terms = new List<string>();
        foreach (var line in skills)
        {
            if (string.IsNullOrWhiteSpace(line) || !line.Contains(':'))
            {
                continue;
            }

            var items = line.Split(':', 2)[1]
                .Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

            foreach (var item in items)
            {
                var term = item.Trim();
                if (term.Length >= 3 || term.Contains('#') || term.Contains('+'))
                {
                    terms.Add(term);
                }
            }
        }

        return terms
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(s => s.Length)
            .ToList();
    }

    private static bool ContainsBoldToken(string text)
    {
        return Regex.IsMatch(text ?? string.Empty, @"\*\*[^*]+\*\*");
    }

    private static bool TryBoldFirstOccurrence(string text, string term, out string updated)
    {
        updated = text;
        if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(term))
        {
            return false;
        }

        var pattern = Regex.Escape(term);
        var regex = new Regex(pattern, RegexOptions.IgnoreCase);
        var match = regex.Match(text);
        while (match.Success)
        {
            var start = match.Index;
            var end = start + match.Length - 1;
            var leftOk = start == 0 || !char.IsLetterOrDigit(text[start - 1]);
            var rightOk = end == text.Length - 1 || !char.IsLetterOrDigit(text[end + 1]);
            if (leftOk && rightOk && !IsInsideBoldSpan(text, start))
            {
                var original = text.Substring(match.Index, match.Length);
                updated = text[..match.Index] + $"**{original}**" + text[(match.Index + match.Length)..];
                return true;
            }

            match = match.NextMatch();
        }

        return false;
    }

    private static bool IsInsideBoldSpan(string text, int index)
    {
        var left = text.LastIndexOf("**", index, StringComparison.Ordinal);
        if (left < 0)
        {
            return false;
        }

        var right = text.IndexOf("**", left + 2, StringComparison.Ordinal);
        return right >= 0 && index > left && index < right;
    }

    private static bool ContainsPersonalData(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var emailRegex = new Regex(@"[\w._%+-]+@[\w.-]+\.[A-Za-z]{2,}");
        var phoneRegex = new Regex(@"\+?\d[\d\s().-]{7,}");
        var urlRegex = new Regex(@"https?://|www\.");

        return emailRegex.IsMatch(text) || phoneRegex.IsMatch(text) || urlRegex.IsMatch(text);
    }

    private static string ComputeYearsOfExperience(Profile profile)
    {
        var earliestYear = profile.Experience
            .Select(exp => ExtractYear(exp.Dates))
            .Where(year => year > 0)
            .DefaultIfEmpty(0)
            .Min();

        if (earliestYear <= 0)
        {
            return "5+";
        }

        var currentYear = DateTime.Now.Year;
        var years = Math.Max(1, currentYear - earliestYear);
        return $"{years}+";
    }

    private static int ExtractYear(string? dates)
    {
        if (string.IsNullOrWhiteSpace(dates))
        {
            return 0;
        }

        var match = Regex.Match(dates, @"(19|20)\d{2}");
        return match.Success ? int.Parse(match.Value) : 0;
    }
}
