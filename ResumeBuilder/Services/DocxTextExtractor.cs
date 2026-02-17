using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ResumeBuilder.Services;

public static class DocxTextExtractor
{
    public static string ExtractText(string path)
    {
        if (!File.Exists(path))
        {
            return string.Empty;
        }

        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var body = doc.MainDocumentPart?.Document.Body;
            if (body == null)
            {
                return string.Empty;
            }

            var texts = body.Descendants<Text>().Select(t => t.Text);
            return string.Join(" ", texts);
        }
        catch
        {
            return string.Empty;
        }
    }
}
