using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SignFill.Models;
using System.IO;

namespace SignFill.Services;

public class WordGenerator
{
    public void Generate(
        string templatePath,
        string outputPath,
        DataRow row)
    {
        if (!File.Exists(templatePath))
        {
            throw new FileNotFoundException(
                "Không tìm thấy Word template.",
                templatePath);
        }

        if (row == null)
        {
            throw new ArgumentNullException(nameof(row));
        }

        File.Copy(templatePath, outputPath, true);

        using WordprocessingDocument document =
            WordprocessingDocument.Open(
                outputPath,
                true);

        if (document.MainDocumentPart == null)
        {
            throw new InvalidDataException(
                "Word document không có MainDocumentPart.");
        }

        var body =
            document.MainDocumentPart.Document?.Body;

        if (body == null)
        {
            throw new InvalidDataException(
                "Word document không có nội dung.");
        }

        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            ReplaceInParagraph(
                paragraph,
                row);
        }

        document.MainDocumentPart.Document.Save();
    }

    private static void ReplaceInParagraph(
        Paragraph paragraph,
        DataRow row)
    {
        var textElements =
            paragraph.Descendants<Text>().ToList();

        if (textElements.Count == 0)
        {
            return;
        }

        string combinedText =
            string.Concat(
                textElements.Select(text => text.Text));

        string replacedText =
            ReplaceFields(
                combinedText,
                row);

        if (combinedText == replacedText)
        {
            return;
        }

        // Đưa toàn bộ kết quả vào Text đầu tiên.
        textElements[0].Text = replacedText;

        // Các Text còn lại được xóa nội dung.
        for (int i = 1;
             i < textElements.Count;
             i++)
        {
            textElements[i].Text = string.Empty;
        }
    }

    private static string ReplaceFields(
        string text,
        DataRow row)
    {
        string result = text;

        foreach (var field in row.Fields)
        {
            string placeholder =
                $"{{{{{field.Key}}}}}";

            result =
                result.Replace(
                    placeholder,
                    field.Value);
        }

        return result;
    }
}