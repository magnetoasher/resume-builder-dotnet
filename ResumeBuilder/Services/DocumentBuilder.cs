using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ResumeBuilder.Models;

namespace ResumeBuilder.Services;

public class DocumentBuilder
{
    private const string DefaultFontFamily = "Calibri";
    private const int DefaultNameFontSize = 28;
    private const int DefaultContactFontSize = 11;
    private const int DefaultHeadingFontSize = 13;
    private const int DefaultBodyFontSize = 11;
    private const int BulletAbstractNumberingId = 1;
    private const int BulletNumberingId = 1;

    private static readonly HashSet<string> SectionNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "summary",
        "professional summary",
        "career summary",
        "profile summary",
        "technical skills",
        "technology skills",
        "relevant skills",
        "core skills",
        "skills",
        "education",
        "professional experience",
        "experience",
        "work experience",
        "employment history",
        "licenses & certifications",
        "certifications"
    };

    public void BuildResumeDocx(ResumeContent content, Profile profile, string outputPath)
    {
        BuildDefault(content, profile, outputPath);
    }

    private static void BuildFromTemplate(string templatePath, ResumeContent content, Profile profile, string outputPath)
    {
        System.IO.File.Copy(templatePath, outputPath, overwrite: true);

        using var doc = WordprocessingDocument.Open(outputPath, true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Invalid template document.");
        EnsureBulletNumbering(mainPart);

        var body = mainPart.Document.Body ?? throw new InvalidOperationException("Invalid template body.");
        var style = TemplateStyleSnapshot.Capture(body);
        var section = body.Elements<SectionProperties>().FirstOrDefault()?.CloneNode(true);
        body.RemoveAllChildren();

        WriteResume(body, content, profile, style);
        if (section != null)
        {
            body.Append(section);
        }

        mainPart.Document.Save();
    }

    private static void BuildDefault(ResumeContent content, Profile profile, string outputPath)
    {
        using var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        EnsureBulletNumbering(mainPart);

        var body = mainPart.Document.Body ?? throw new InvalidOperationException("Could not create document body.");
        var style = TemplateStyleSnapshot.Default();
        WriteResume(body, content, profile, style);
        body.Append(CreateDefaultSectionProperties());
        mainPart.Document.Save();
    }

    private static SectionProperties CreateDefaultSectionProperties()
    {
        return new SectionProperties(
            new PageSize
            {
                Width = 12240U,
                Height = 15840U
            },
            new PageMargin
            {
                Top = 720,
                Right = 720U,
                Bottom = 720,
                Left = 720U,
                Header = 720U,
                Footer = 720U,
                Gutter = 0U
            });
    }

    private static void WriteResume(Body body, ResumeContent content, Profile profile, TemplateStyleSnapshot style)
    {
        AppendParagraph(body, profile.DisplayName, DefaultNameFontSize, bold: true, style.Name, centerOverride: true);

        var contact = ComposeContactLine(profile);
        if (!string.IsNullOrWhiteSpace(contact))
        {
            AppendParagraph(body, contact, DefaultContactFontSize, bold: false, style.Contact, centerOverride: true);
        }

        AppendSpacer(body);

        AppendParagraph(body, style.SummaryHeadingText, DefaultHeadingFontSize, bold: true, style.SummaryHeading);
        foreach (var block in SplitBlocks(content.Summary))
        {
            AppendParagraph(body, block, DefaultBodyFontSize, bold: false, style.SummaryBody, centerOverride: false);
        }

        AppendSpacer(body);

        if (style.SkillsBeforeEducation)
        {
            AppendSkillsSection(body, content, style);
            AppendEducationSection(body, content, style);
        }
        else
        {
            AppendEducationSection(body, content, style);
            AppendSkillsSection(body, content, style);
        }

        AppendParagraph(body, style.ExperienceHeadingText, DefaultHeadingFontSize, bold: true, style.ExperienceHeading);
        foreach (var role in content.Experience)
        {
            var left = ComposeRoleHeaderLeft(role);
            AppendLeftRightParagraph(
                body,
                left,
                role.Dates?.Trim() ?? string.Empty,
                DefaultBodyFontSize,
                leftBold: true,
                rightBold: true,
                style.ExperienceHeader,
                rightTabPosition: style.RightAlignedTabPosition,
                centerOverride: false);

            foreach (var bullet in role.Bullets.Where(b => !string.IsNullOrWhiteSpace(b)))
            {
                AppendParagraph(body, CleanSentence(bullet), DefaultBodyFontSize, bold: false, style.ExperienceBullet, centerOverride: false, bullet: true);
            }

            AppendSpacer(body);
        }
    }

    private static void AppendSkillsSection(Body body, ResumeContent content, TemplateStyleSnapshot style)
    {
        if (content.Skills.Count == 0)
        {
            return;
        }

        AppendParagraph(body, style.SkillsHeadingText, DefaultHeadingFontSize, bold: true, style.SkillsHeading);

        if (style.SkillsAsBullets)
        {
            var items = FlattenSkillItems(content.Skills).Take(24).ToList();
            AppendSkillColumns(body, items, style.SkillsBody);
        }
        else
        {
            foreach (var line in content.Skills.Where(s => !string.IsNullOrWhiteSpace(s)))
            {
                AppendSkillLine(body, line.Trim(), style.SkillsBody, centerOverride: false);
            }
        }

        AppendSpacer(body);
    }

    private static void AppendSkillColumns(Body body, List<string> items, ParagraphSnapshot style)
    {
        if (items.Count == 0)
        {
            return;
        }

        if (items.Count < 6)
        {
            foreach (var item in items)
            {
                AppendParagraph(body, item, DefaultBodyFontSize, bold: false, style, centerOverride: false, bullet: true);
            }

            return;
        }

        const int columns = 3;
        var groups = SplitIntoColumns(items, columns);
        var rows = groups.Max(g => g.Count);

        var table = new Table();
        var tableProps = new TableProperties(
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
            new TableBorders(
                new TopBorder { Val = BorderValues.None },
                new LeftBorder { Val = BorderValues.None },
                new BottomBorder { Val = BorderValues.None },
                new RightBorder { Val = BorderValues.None },
                new InsideHorizontalBorder { Val = BorderValues.None },
                new InsideVerticalBorder { Val = BorderValues.None }));
        table.Append(tableProps);

        for (var r = 0; r < rows; r++)
        {
            var row = new TableRow();
            for (var c = 0; c < columns; c++)
            {
                var cell = new TableCell();
                var text = r < groups[c].Count ? groups[c][r] : string.Empty;
                var paragraph = CreateParagraphShell(style, centerOverride: false, bullet: false);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    paragraph.Append(CreateRun("\u2022 ", DefaultBodyFontSize, style, bold: false));
                    AppendRunsWithInlineBold(paragraph, text, DefaultBodyFontSize, style, bold: false);
                }

                cell.Append(paragraph);
                cell.Append(new TableCellProperties(new TableCellWidth
                {
                    Type = TableWidthUnitValues.Pct,
                    Width = "1667"
                }));
                row.Append(cell);
            }

            table.Append(row);
        }

        body.Append(table);
    }

    private static List<List<string>> SplitIntoColumns(List<string> items, int columns)
    {
        var groups = Enumerable.Range(0, columns).Select(_ => new List<string>()).ToList();
        var perColumn = (int)Math.Ceiling(items.Count / (double)columns);
        var index = 0;
        for (var c = 0; c < columns; c++)
        {
            for (var i = 0; i < perColumn && index < items.Count; i++)
            {
                groups[c].Add(items[index]);
                index++;
            }
        }

        return groups;
    }

    private static void AppendEducationSection(Body body, ResumeContent content, TemplateStyleSnapshot style)
    {
        AppendParagraph(body, style.EducationHeadingText, DefaultHeadingFontSize, bold: true, style.EducationHeading);

        if (content.Education.Count == 0)
        {
            AppendSpacer(body);
            return;
        }

        for (var i = 0; i < content.Education.Count; i++)
        {
            var entry = content.Education[i];
            var school = entry.School?.Trim() ?? string.Empty;
            var degree = entry.Degree?.Trim() ?? string.Empty;
            var dates = entry.Dates?.Trim() ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(school) || !string.IsNullOrWhiteSpace(dates))
            {
                AppendLeftRightParagraph(
                    body,
                    school,
                    dates,
                    DefaultBodyFontSize,
                    leftBold: false,
                    rightBold: true,
                    style.EducationBody,
                    rightTabPosition: style.RightAlignedTabPosition,
                    centerOverride: false);
            }

            if (!string.IsNullOrWhiteSpace(degree))
            {
                AppendParagraph(body, degree, DefaultBodyFontSize, bold: true, style.EducationBody, centerOverride: false);
            }

            if (i < content.Education.Count - 1)
            {
                AppendSpacer(body);
            }
        }

        AppendSpacer(body);
    }

    private static List<string> FlattenSkillItems(IEnumerable<string> skillLines)
    {
        var items = new List<string>();
        foreach (var line in skillLines)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            var split = line.Split(':', 2);
            var payload = split.Length == 2 ? split[1] : split[0];
            foreach (var raw in payload.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                var item = raw.Trim();
                if (!string.IsNullOrWhiteSpace(item))
                {
                    items.Add(item);
                }
            }
        }

        return items
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string ComposeRoleHeaderLeft(ResumeExperience role)
    {
        var title = role.Title?.Trim() ?? string.Empty;
        var company = role.Company?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(title))
        {
            return company;
        }

        if (string.IsNullOrWhiteSpace(company))
        {
            return title;
        }

        return $"{title} | {company}";
    }

    private static string ComposeContactLine(Profile profile)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(profile.Address))
        {
            parts.Add(profile.Address.Trim());
        }

        if (!string.IsNullOrWhiteSpace(profile.Email))
        {
            parts.Add(profile.Email.Trim());
        }

        if (!string.IsNullOrWhiteSpace(profile.Phone))
        {
            parts.Add(profile.Phone.Trim());
        }

        if (!string.IsNullOrWhiteSpace(profile.LinkedIn))
        {
            parts.Add(CompactLinkedIn(profile.LinkedIn));
        }

        if (parts.Count > 0)
        {
            return string.Join(" \u2022 ", parts);
        }

        return profile.ContactLine?.Trim() ?? string.Empty;
    }

    private static string CompactLinkedIn(string value)
    {
        var link = (value ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(link))
        {
            return string.Empty;
        }

        if (Uri.TryCreate(link, UriKind.Absolute, out var uri))
        {
            var host = uri.Host.StartsWith("www.", StringComparison.OrdinalIgnoreCase)
                ? uri.Host[4..]
                : uri.Host;
            var path = uri.AbsolutePath.TrimEnd('/');
            return string.IsNullOrWhiteSpace(path) ? host : $"{host}{path}";
        }

        return link
            .Replace("https://", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("http://", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("www.", string.Empty, StringComparison.OrdinalIgnoreCase)
            .TrimEnd('/');
    }

    private static IEnumerable<string> SplitBlocks(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            yield break;
        }

        foreach (var line in text.Split('\n'))
        {
            var block = line.Trim();
            if (!string.IsNullOrWhiteSpace(block))
            {
                yield return block;
            }
        }
    }

    private static string FormatEducationLine(Education entry)
    {
        var degree = entry.Degree?.Trim() ?? string.Empty;
        var school = entry.School?.Trim() ?? string.Empty;
        var dates = entry.Dates?.Trim() ?? string.Empty;

        var line = degree;
        if (!string.IsNullOrWhiteSpace(school))
        {
            line = string.IsNullOrWhiteSpace(line) ? school : $"{line} | {school}";
        }

        if (!string.IsNullOrWhiteSpace(dates))
        {
            line = string.IsNullOrWhiteSpace(line) ? dates : $"{line}    {dates}";
        }

        return line;
    }

    private static void AppendSkillLine(Body body, string line, ParagraphSnapshot style, bool? centerOverride)
    {
        var colon = line.IndexOf(':');
        if (colon < 0)
        {
            AppendParagraph(body, line, DefaultBodyFontSize, bold: false, style, centerOverride: centerOverride);
            return;
        }

        var category = line[..(colon + 1)].TrimEnd();
        var value = line[(colon + 1)..].TrimStart();

        var paragraph = CreateParagraphShell(style, centerOverride, bullet: false);
        paragraph.Append(CreateRun(category, DefaultBodyFontSize, style, bold: true));
        if (!string.IsNullOrWhiteSpace(value))
        {
            paragraph.Append(CreateRun($" {value}", DefaultBodyFontSize, style, bold: false));
        }

        body.Append(paragraph);
    }

    private static void AppendExperienceHeader(Body body, string company, string dates, ParagraphSnapshot style, bool? centerOverride)
    {
        var paragraph = CreateParagraphShell(style, centerOverride, bullet: false);
        paragraph.Append(CreateRun(company?.Trim() ?? string.Empty, DefaultBodyFontSize, style, bold: true));
        if (!string.IsNullOrWhiteSpace(dates))
        {
            paragraph.Append(CreateRun($" | {dates.Trim()}", DefaultBodyFontSize, style, bold: false));
        }

        body.Append(paragraph);
    }

    private static void AppendParagraph(
        Body body,
        string text,
        int fallbackFontSize,
        bool bold,
        ParagraphSnapshot style,
        bool? centerOverride = null,
        bool bullet = false)
    {
        var paragraph = CreateParagraphShell(style, centerOverride, bullet);
        var value = text ?? string.Empty;

        if (bullet && style.Numbering == null)
        {
            paragraph.Append(CreateRun("\u2022\t", fallbackFontSize, style, bold: false));
            AppendRunsWithInlineBold(paragraph, value, fallbackFontSize, style, bold);
            body.Append(paragraph);
            return;
        }

        AppendRunsWithInlineBold(paragraph, value, fallbackFontSize, style, bold);

        body.Append(paragraph);
    }

    private static void AppendLeftRightParagraph(
        Body body,
        string leftText,
        string rightText,
        int fallbackFontSize,
        bool leftBold,
        bool rightBold,
        ParagraphSnapshot style,
        int? rightTabPosition = null,
        bool? centerOverride = false)
    {
        var paragraph = CreateParagraphShell(style, centerOverride, bullet: false);
        var left = leftText?.Trim() ?? string.Empty;
        var right = rightText?.Trim() ?? string.Empty;
        var hasBothSides = !string.IsNullOrWhiteSpace(left) && !string.IsNullOrWhiteSpace(right);

        if (hasBothSides)
        {
            EnsureRightTabStop(paragraph, rightTabPosition ?? ResolveRightTabPosition(style));
        }

        if (!string.IsNullOrWhiteSpace(left))
        {
            AppendRunsWithInlineBold(paragraph, left, fallbackFontSize, style, leftBold);
        }

        if (!string.IsNullOrWhiteSpace(right))
        {
            if (!string.IsNullOrWhiteSpace(left))
            {
                paragraph.Append(CreateRun("\t", fallbackFontSize, style, bold: false));
            }

            AppendRunsWithInlineBold(paragraph, right, fallbackFontSize, style, rightBold);
        }

        body.Append(paragraph);
    }

    private static int ResolveRightTabPosition(ParagraphSnapshot style)
    {
        var tabs = style.ParagraphProps?.Tabs?.Elements<TabStop>().ToList();
        if (tabs == null || tabs.Count == 0)
        {
            return 9360;
        }

        List<int> rightStops = tabs
            .Where(t => t.Val?.Value == TabStopValues.Right && t.Position?.Value != null)
            .Select(t => (int)t.Position!.Value!)
            .ToList();

        if (rightStops.Count > 0)
        {
            return rightStops.Max();
        }

        List<int> allStops = tabs
            .Where(t => t.Position?.Value != null)
            .Select(t => (int)t.Position!.Value!)
            .ToList();

        return allStops.Count > 0 ? allStops.Max() : 9360;
    }

    private static void EnsureRightTabStop(Paragraph paragraph, int position)
    {
        paragraph.ParagraphProperties ??= new ParagraphProperties();
        var props = paragraph.ParagraphProperties;

        props.RemoveAllChildren<Tabs>();
        var tabs = new Tabs(
            new TabStop
            {
                Val = TabStopValues.Right,
                Position = position
            });
        props.Append(tabs);
    }

    private static void AppendRunsWithInlineBold(
        Paragraph paragraph,
        string value,
        int fallbackFontSize,
        ParagraphSnapshot style,
        bool bold)
    {
        if (!value.Contains("**", StringComparison.Ordinal))
        {
            paragraph.Append(CreateRun(value, fallbackFontSize, style, bold));
            return;
        }

        foreach (var part in Regex.Split(value, "(\\*\\*.*?\\*\\*)"))
        {
            if (string.IsNullOrEmpty(part))
            {
                continue;
            }

            if (part.StartsWith("**", StringComparison.Ordinal) && part.EndsWith("**", StringComparison.Ordinal) && part.Length > 4)
            {
                paragraph.Append(CreateRun(part[2..^2], fallbackFontSize, style, bold: true));
            }
            else
            {
                paragraph.Append(CreateRun(part, fallbackFontSize, style, bold));
            }
        }
    }

    private static Paragraph CreateParagraphShell(ParagraphSnapshot style, bool? centerOverride, bool bullet)
    {
        var paragraph = new Paragraph();
        var props = style.ParagraphProps != null
            ? (ParagraphProperties)style.ParagraphProps.CloneNode(true)
            : new ParagraphProperties();

        // Body-level section changes from the template should not be copied per paragraph.
        props.RemoveAllChildren<SectionProperties>();

        props.RemoveAllChildren<ParagraphStyleId>();
        if (!string.IsNullOrWhiteSpace(style.StyleId))
        {
            props.Append(new ParagraphStyleId { Val = style.StyleId });
        }

        props.RemoveAllChildren<NumberingProperties>();
        if (bullet)
        {
            if (style.Numbering != null)
            {
                props.Append((NumberingProperties)style.Numbering.CloneNode(true));
            }
            else
            {
                props.RemoveAllChildren<Indentation>();
                props.Append(new Indentation
                {
                    Left = "720",
                    Hanging = "360"
                });

                // Align wrapped lines under the first character after the bullet.
                props.RemoveAllChildren<Tabs>();
                props.Append(new Tabs(
                    new TabStop
                    {
                        Val = TabStopValues.Left,
                        Position = 720
                    }));
            }
        }

        var effectiveJustification = style.Justification;
        var useCenter = centerOverride ?? (effectiveJustification == JustificationValues.Center);
        props.RemoveAllChildren<Justification>();
        if (useCenter)
        {
            props.Append(new Justification { Val = JustificationValues.Center });
        }
        else if (centerOverride.HasValue && effectiveJustification == JustificationValues.Center)
        {
            props.Append(new Justification { Val = JustificationValues.Left });
        }
        else if (effectiveJustification.HasValue)
        {
            props.Append(new Justification { Val = effectiveJustification.Value });
        }

        if (props.ChildElements.Count > 0)
        {
            paragraph.ParagraphProperties = props;
        }

        return paragraph;
    }

    private static Run CreateRun(string text, int fallbackFontSize, ParagraphSnapshot style, bool bold)
    {
        RunProperties runProps;
        if (style.RunProps != null)
        {
            runProps = (RunProperties)style.RunProps.CloneNode(true);
        }
        else
        {
            runProps = new RunProperties();
        }

        NormalizeRunProperties(runProps, fallbackFontSize);

        if (!bold)
        {
            runProps.RemoveAllChildren<Bold>();
        }
        else if (!runProps.Elements<Bold>().Any())
        {
            runProps.Append(new Bold());
        }

        var run = new Run();
        if (runProps.ChildElements.Count > 0)
        {
            run.Append(runProps);
        }

        run.Append(new Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static void NormalizeRunProperties(RunProperties runProps, int fontSize)
    {
        var fonts = runProps.GetFirstChild<RunFonts>();
        if (fonts == null)
        {
            fonts = new RunFonts();
            runProps.PrependChild(fonts);
        }

        fonts.Ascii ??= DefaultFontFamily;
        fonts.HighAnsi ??= DefaultFontFamily;
        fonts.ComplexScript ??= DefaultFontFamily;

        if (!runProps.Elements<FontSize>().Any())
        {
            runProps.Append(new FontSize { Val = (fontSize * 2).ToString() });
        }

        if (!runProps.Elements<FontSizeComplexScript>().Any())
        {
            runProps.Append(new FontSizeComplexScript { Val = (fontSize * 2).ToString() });
        }
    }

    private static void EnsureBulletNumbering(MainDocumentPart mainPart)
    {
        var numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering ??= new Numbering();

        var hasNumberingInstance = numberingPart.Numbering.Elements<NumberingInstance>()
            .Any(n => n.NumberID?.Value == BulletNumberingId);
        if (hasNumberingInstance)
        {
            return;
        }

        var abstractNumbering = new AbstractNum { AbstractNumberId = BulletAbstractNumberingId };
        var level = new Level { LevelIndex = 0 };
        level.Append(new StartNumberingValue { Val = 1 });
        level.Append(new NumberingFormat { Val = NumberFormatValues.Bullet });
        level.Append(new LevelText { Val = "\u2022" });
        level.Append(new LevelJustification { Val = LevelJustificationValues.Left });
        level.Append(new PreviousParagraphProperties(new Indentation { Left = "720", Hanging = "360" }));
        level.Append(new NumberingSymbolRunProperties(new RunFonts { Ascii = "Symbol", HighAnsi = "Symbol" }));
        abstractNumbering.Append(level);

        var instance = new NumberingInstance { NumberID = BulletNumberingId };
        instance.Append(new AbstractNumId { Val = BulletAbstractNumberingId });

        numberingPart.Numbering.Append(abstractNumbering);
        numberingPart.Numbering.Append(instance);
        numberingPart.Numbering.Save();
    }

    private static string CleanSentence(string text)
    {
        var cleaned = Regex.Replace(text ?? string.Empty, @"^[\s\u2022\u2023\u25E6\u2043\u2219\u00B7\uF0B7\-]+", string.Empty).Trim();
        cleaned = Regex.Replace(cleaned, @"\s+", " ");
        if (!string.IsNullOrWhiteSpace(cleaned) && !Regex.IsMatch(cleaned, @"[.!?]$"))
        {
            cleaned += ".";
        }

        return cleaned;
    }

    private static void AppendSpacer(Body body)
    {
        var paragraph = new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines
                {
                    Before = "80",
                    After = "180",
                    Line = "280",
                    LineRule = LineSpacingRuleValues.Auto
                }),
            new Run(new Text(string.Empty)));
        body.Append(paragraph);
    }

    private sealed class TemplateStyleSnapshot
    {
        public ParagraphSnapshot Name { get; init; } = ParagraphSnapshot.Empty;
        public ParagraphSnapshot Contact { get; init; } = ParagraphSnapshot.Empty;
        public ParagraphSnapshot SummaryHeading { get; init; } = ParagraphSnapshot.Empty;
        public string SummaryHeadingText { get; init; } = "Summary:";
        public ParagraphSnapshot SummaryBody { get; init; } = ParagraphSnapshot.Empty;
        public ParagraphSnapshot SkillsHeading { get; init; } = ParagraphSnapshot.Empty;
        public string SkillsHeadingText { get; init; } = "Technical Skills:";
        public ParagraphSnapshot SkillsBody { get; init; } = ParagraphSnapshot.Empty;
        public ParagraphSnapshot EducationHeading { get; init; } = ParagraphSnapshot.Empty;
        public string EducationHeadingText { get; init; } = "Education:";
        public ParagraphSnapshot EducationBody { get; init; } = ParagraphSnapshot.Empty;
        public ParagraphSnapshot ExperienceHeading { get; init; } = ParagraphSnapshot.Empty;
        public string ExperienceHeadingText { get; init; } = "Professional Experience:";
        public ParagraphSnapshot ExperienceHeader { get; init; } = ParagraphSnapshot.Empty;
        public ParagraphSnapshot ExperienceTitle { get; init; } = ParagraphSnapshot.Empty;
        public ParagraphSnapshot ExperienceBullet { get; init; } = ParagraphSnapshot.Empty;
        public int RightAlignedTabPosition { get; init; } = 10800;
        public bool SkillsAsBullets { get; init; }
        public bool SkillsBeforeEducation { get; init; }

        public static TemplateStyleSnapshot Default()
        {
            var name = CreateTextSnapshot(JustificationValues.Center, "Times New Roman", 32, bold: true, after: "120", line: "300");
            var contact = CreateTextSnapshot(JustificationValues.Center, "Times New Roman", 20, bold: false, after: "120", line: "280");
            var heading = CreateTextSnapshot(JustificationValues.Left, "Times New Roman", 24, bold: true, after: "100", line: "280");
            var body = CreateTextSnapshot(JustificationValues.Left, "Times New Roman", 24, bold: false, after: "40", line: "280");
            var bodyBold = CreateTextSnapshot(JustificationValues.Left, "Times New Roman", 24, bold: true, after: "40", line: "280");

            return new TemplateStyleSnapshot
            {
                Name = name,
                Contact = contact,
                SummaryHeading = heading,
                SummaryHeadingText = "CAREER SUMMARY:",
                SummaryBody = body,
                SkillsHeading = heading,
                SkillsHeadingText = "TECHNOLOGY SKILLS:",
                SkillsBody = body,
                EducationHeading = heading,
                EducationHeadingText = "EDUCATION:",
                EducationBody = bodyBold,
                ExperienceHeading = heading,
                ExperienceHeadingText = "PROFESSIONAL EXPERIENCE:",
                ExperienceHeader = bodyBold,
                ExperienceTitle = body,
                ExperienceBullet = body,
                RightAlignedTabPosition = 10800,
                SkillsAsBullets = true,
                SkillsBeforeEducation = true
            };
        }

        public static TemplateStyleSnapshot Capture(Body body)
        {
            var paragraphs = body.Elements<Paragraph>().ToList();
            var nonEmpty = paragraphs.Where(p => !string.IsNullOrWhiteSpace(p.InnerText)).ToList();

            var summaryHeading = FindHeading(nonEmpty, "career summary", "professional summary", "summary", "profile summary");
            var skillsHeading = FindHeading(nonEmpty, "technology skills", "technical skills", "core skills", "skills", "relevant skills");
            var educationHeading = FindHeading(nonEmpty, "education");
            var experienceHeading = FindHeading(nonEmpty, "professional experience", "experience", "work experience", "employment history");

            var summaryBody = FindNextContent(nonEmpty, summaryHeading, allowBullets: false);
            var skillsBodyAny = FindNextContent(nonEmpty, skillsHeading, allowBullets: true);
            var skillsBodyText = FindNextContent(nonEmpty, skillsHeading, allowBullets: false);
            var educationBody = FindNextContent(nonEmpty, educationHeading, allowBullets: false);

            var expHeader = FindNextContent(nonEmpty, experienceHeading, allowBullets: false);
            var expTitle = FindNextAfter(nonEmpty, expHeader, p => !IsSectionHeading(p) && !IsBulletParagraph(p));
            var expBullet = FindNextAfter(nonEmpty, experienceHeading, IsBulletParagraph) ?? nonEmpty.FirstOrDefault(IsBulletParagraph);

            var skillsIndex = skillsHeading != null ? nonEmpty.IndexOf(skillsHeading) : -1;
            var educationIndex = educationHeading != null ? nonEmpty.IndexOf(educationHeading) : -1;
            var skillsBeforeEducation = skillsIndex >= 0 && educationIndex >= 0 && skillsIndex < educationIndex;
            var skillsAsBullets = (skillsBodyAny?.ParagraphProperties?.NumberingProperties != null)
                || (skillsBodyAny != null && IsBulletParagraph(skillsBodyAny));
            var skillsBody = skillsAsBullets ? skillsBodyAny : skillsBodyText;
            var rightAlignedTabPosition = ResolveBodyRightEdge(body);

            return new TemplateStyleSnapshot
            {
                Name = ParagraphSnapshot.FromParagraph(nonEmpty.ElementAtOrDefault(0)),
                Contact = ParagraphSnapshot.FromParagraph(nonEmpty.ElementAtOrDefault(1)),
                SummaryHeading = ParagraphSnapshot.FromParagraph(summaryHeading),
                SummaryHeadingText = GetHeadingText(summaryHeading, "Summary:"),
                SummaryBody = ParagraphSnapshot.FromParagraph(summaryBody),
                SkillsHeading = ParagraphSnapshot.FromParagraph(skillsHeading),
                SkillsHeadingText = GetHeadingText(skillsHeading, "Technical Skills:"),
                SkillsBody = ParagraphSnapshot.FromParagraph(skillsBody),
                EducationHeading = ParagraphSnapshot.FromParagraph(educationHeading),
                EducationHeadingText = GetHeadingText(educationHeading, "Education:"),
                EducationBody = ParagraphSnapshot.FromParagraph(educationBody),
                ExperienceHeading = ParagraphSnapshot.FromParagraph(experienceHeading),
                ExperienceHeadingText = GetHeadingText(experienceHeading, "Professional Experience:"),
                ExperienceHeader = ParagraphSnapshot.FromParagraph(expHeader),
                ExperienceTitle = ParagraphSnapshot.FromParagraph(expTitle),
                ExperienceBullet = ParagraphSnapshot.FromParagraph(expBullet),
                RightAlignedTabPosition = rightAlignedTabPosition,
                SkillsAsBullets = skillsAsBullets,
                SkillsBeforeEducation = skillsBeforeEducation
            };
        }

        private static int ResolveBodyRightEdge(Body body)
        {
            const int fallback = 10800;
            var section = body.Elements<SectionProperties>().LastOrDefault();
            if (section == null)
            {
                return fallback;
            }

            var pageWidth = section.GetFirstChild<PageSize>()?.Width?.Value;
            var margins = section.GetFirstChild<PageMargin>();
            var leftMargin = margins?.Left?.Value;
            var rightMargin = margins?.Right?.Value;

            if (!pageWidth.HasValue || !leftMargin.HasValue || !rightMargin.HasValue)
            {
                return fallback;
            }

            var width = (int)pageWidth.Value - (int)leftMargin.Value - (int)rightMargin.Value;
            return width > 0 ? width : fallback;
        }

        private static ParagraphSnapshot CreateTextSnapshot(
            JustificationValues justification,
            string font,
            int fontHalfPoints,
            bool bold,
            string after = "0",
            string line = "240")
        {
            var paragraphProps = new ParagraphProperties(
                new SpacingBetweenLines
                {
                    After = after,
                    Line = line,
                    LineRule = LineSpacingRuleValues.Auto
                },
                new Justification { Val = justification });

            var runProps = new RunProperties(
                new RunFonts
                {
                    Ascii = font,
                    HighAnsi = font,
                    ComplexScript = font
                },
                new FontSize { Val = fontHalfPoints.ToString() },
                new FontSizeComplexScript { Val = fontHalfPoints.ToString() });

            if (bold)
            {
                runProps.Append(new Bold());
            }

            return new ParagraphSnapshot
            {
                Justification = justification,
                ParagraphProps = paragraphProps,
                RunProps = runProps
            };
        }

        private static Paragraph? FindHeading(IEnumerable<Paragraph> paragraphs, params string[] aliases)
        {
            foreach (var alias in aliases)
            {
                var normalized = NormalizeHeading(alias);
                var match = paragraphs.FirstOrDefault(p => NormalizeHeading(p.InnerText) == normalized);
                if (match != null)
                {
                    return match;
                }
            }

            return null;
        }

        private static string GetHeadingText(Paragraph? paragraph, string fallback)
        {
            var text = paragraph?.InnerText?.Trim();
            return string.IsNullOrWhiteSpace(text) ? fallback : text;
        }

        private static Paragraph? FindNextContent(List<Paragraph> paragraphs, Paragraph? anchor, bool allowBullets)
        {
            if (anchor == null)
            {
                return null;
            }

            return FindNextAfter(
                paragraphs,
                anchor,
                p => !IsSectionHeading(p) && (allowBullets || !IsBulletParagraph(p)));
        }

        private static Paragraph? FindNextAfter(List<Paragraph> paragraphs, Paragraph? anchor, Func<Paragraph, bool> predicate)
        {
            if (anchor == null)
            {
                return null;
            }

            var index = paragraphs.IndexOf(anchor);
            if (index < 0)
            {
                return paragraphs.FirstOrDefault(predicate);
            }

            for (var i = index + 1; i < paragraphs.Count; i++)
            {
                var paragraph = paragraphs[i];
                if (predicate(paragraph))
                {
                    return paragraph;
                }
            }

            return null;
        }

        private static bool IsSectionHeading(Paragraph paragraph)
        {
            var heading = NormalizeHeading(paragraph.InnerText);
            return SectionNames.Contains(heading);
        }

        private static bool IsBulletParagraph(Paragraph paragraph)
        {
            if (paragraph.ParagraphProperties?.NumberingProperties != null)
            {
                return true;
            }

            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? string.Empty;
            if (styleId.IndexOf("List", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            var text = paragraph.InnerText?.TrimStart() ?? string.Empty;
            return text.StartsWith("\u2022", StringComparison.Ordinal)
                || text.StartsWith("-", StringComparison.Ordinal)
                || text.StartsWith("\uF0B7", StringComparison.Ordinal);
        }

        private static string NormalizeHeading(string text)
        {
            var value = (text ?? string.Empty).Trim().TrimEnd(':').Trim();
            value = Regex.Replace(value, @"\s+", " ");
            return value.ToLowerInvariant();
        }
    }

    private sealed class ParagraphSnapshot
    {
        public static ParagraphSnapshot Empty { get; } = new();

        public string? StyleId { get; init; }
        public JustificationValues? Justification { get; init; }
        public ParagraphProperties? ParagraphProps { get; init; }
        public RunProperties? RunProps { get; init; }
        public NumberingProperties? Numbering { get; init; }

        public static ParagraphSnapshot FromParagraph(Paragraph? paragraph)
        {
            if (paragraph == null)
            {
                return Empty;
            }

            var paragraphProps = paragraph.ParagraphProperties;
            var styleId = paragraphProps?.ParagraphStyleId?.Val?.Value;
            var justification = paragraphProps?.Justification?.Val?.Value;
            var runProps = paragraph.Descendants<RunProperties>().FirstOrDefault();
            var numbering = paragraphProps?.NumberingProperties;

            return new ParagraphSnapshot
            {
                StyleId = styleId,
                Justification = justification,
                ParagraphProps = paragraphProps != null ? (ParagraphProperties)paragraphProps.CloneNode(true) : null,
                RunProps = runProps != null ? (RunProperties)runProps.CloneNode(true) : null,
                Numbering = numbering != null ? (NumberingProperties)numbering.CloneNode(true) : null
            };
        }
    }
}
