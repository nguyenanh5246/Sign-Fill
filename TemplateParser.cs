using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Text.RegularExpressions;

namespace SignFill.Services;

public class TemplateParser
{
    private static readonly Regex FieldRegex =
        new(@"\{\{(.*?)\}\}",
            RegexOptions.Compiled);

    public List<string> GetFields(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException(
                "Không tìm thấy file Word template.",
                filePath);
        }

        using WordprocessingDocument document =
            WordprocessingDocument.Open(
                filePath,
                false);

        if (document.MainDocumentPart == null)
        {
            throw new InvalidDataException(
                "Word template không có MainDocumentPart.");
        }

        string documentText =
            GetDocumentText(document);

        MatchCollection matches =
            FieldRegex.Matches(documentText);

        List<string> fields = new();

        foreach (Match match in matches)
        {
            string fieldName =
                match.Groups[1]
                     .Value
                     .Trim();

            if (string.IsNullOrWhiteSpace(fieldName))
            {
                continue;
            }

            if (!fields.Contains(
                    fieldName,
                    StringComparer.Ordinal))
            {
                fields.Add(fieldName);
            }
        }

        return fields;
    }

    private static string GetDocumentText(
        WordprocessingDocument document)
    {
        var body =
            document.MainDocumentPart?
                    .Document?
                    .Body;

        if (body == null)
        {
            throw new InvalidDataException(
                "Word template không có nội dung.");
        }

        return string.Join(
            Environment.NewLine,
            body.Descendants<Text>()
                .Select(text => text.Text));
    }
}