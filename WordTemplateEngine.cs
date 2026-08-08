using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Sign_Fill;

public class WordTemplateEngine
{
    public void Generate(
        string templatePath,
        string outputPath,
        List<string> names)
    {
        if (!File.Exists(templatePath))
        {
            throw new FileNotFoundException(
                "Không tìm thấy file Word template.",
                templatePath);
        }

        if (names == null || names.Count == 0)
        {
            throw new ArgumentException(
                "Danh sách tên đang trống.",
                nameof(names));
        }

        // Tạo bản sao từ template
        File.Copy(templatePath, outputPath, true);

        using WordprocessingDocument document =
            WordprocessingDocument.Open(outputPath, true);

        if (document.MainDocumentPart == null ||
            document.MainDocumentPart.Document == null ||
            document.MainDocumentPart.Document.Body == null)
        {
            throw new InvalidDataException(
                "File Word không có MainDocumentPart hoặc Body.");
        }

        var body = document.MainDocumentPart.Document.Body;

        // Thay placeholder trong toàn bộ paragraph (bao gồm cả paragraph nằm trong table)
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            ReplaceInParagraph(paragraph, names);
        }

        // Lưu tài liệu
        document.MainDocumentPart.Document.Save();
    }


    private void ReplaceInParagraph(
        Paragraph paragraph,
        List<string> names)
    {
        foreach (var run in paragraph.Elements<Run>())
        {
            var text = run.GetFirstChild<Text>();

            if (text == null)
                continue;

            for (int i = 0; i < names.Count; i++)
            {
                string placeholder =
                    $"{{{{SIGNER_{i + 1}}}}}";

                if (text.Text.Contains(placeholder))
                {
                    text.Text = text.Text.Replace(
                        placeholder,
                        names[i]);
                }
            }
        }
    }
}