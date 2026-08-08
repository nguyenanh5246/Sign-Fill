using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SignFill;

public class WordTemplateEngine
{
    public void Generate(
        string templatePath,
        string outputPath,
        List<string> names)
    {
        File.Copy(templatePath, outputPath, true);

        using WordprocessingDocument document =
            WordprocessingDocument.Open(outputPath, true);

        if (document.MainDocumentPart == null || document.MainDocumentPart.Document?.Body == null)
            throw new InvalidDataException("The Word document is missing a main document part or body.");

        var body = document.MainDocumentPart.Document.Body;

        int index = 1;

        foreach (var paragraph in body.Elements<Paragraph>())
        {
            foreach (var run in paragraph.Elements<Run>())
            {
                var text = run.GetFirstChild<Text>();

                if (text == null)
                    continue;

                string placeholder = $"{{{{SIGNER_{index}}}}}";

                if (text.Text.Contains(placeholder))
                {
                    text.Text = text.Text.Replace(
                        placeholder,
                        names[index - 1]
                    );

                    index++;

                    if (index > names.Count)
                        return;
                }
            }
        }

        document.MainDocumentPart.Document.Save();
    }
}