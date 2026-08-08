using System.IO;
using System.Text;
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

        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            ReplaceInParagraph(paragraph, names);
        }

        document.MainDocumentPart.Document.Save();
    }


    private void ReplaceInParagraph(
        Paragraph paragraph,
        List<string> names)
    {
        var runs = paragraph.Elements<Run>().ToList();

        if (runs.Count == 0)
            return;

        var texts = new List<Text?>();

        foreach (var run in runs)
        {
            texts.Add(run.GetFirstChild<Text>());
        }

        for (int nameIndex = 0;
             nameIndex < names.Count;
             nameIndex++)
        {
            string placeholder =
                $"{{{{SIGNER_{nameIndex + 1}}}}}";

            int startRun = -1;
            int startChar = -1;

            int currentPosition = 0;

            StringBuilder combinedText = new();

            foreach (var text in texts)
            {
                if (text != null)
                {
                    combinedText.Append(text.Text);
                }
            }

            string fullText = combinedText.ToString();

            int placeholderPosition =
                fullText.IndexOf(placeholder,
                    StringComparison.Ordinal);

            if (placeholderPosition < 0)
                continue;

            // Tìm Run chứa ký tự đầu tiên của placeholder
            int position = 0;

            for (int i = 0; i < texts.Count; i++)
            {
                string runText = texts[i]?.Text ?? "";

                if (placeholderPosition >= position &&
                    placeholderPosition < position + runText.Length)
                {
                    startRun = i;
                    startChar = placeholderPosition - position;
                    break;
                }

                position += runText.Length;
            }

            if (startRun < 0)
                continue;

            int remaining =
                placeholder.Length;

            // Vị trí bắt đầu của placeholder
            int runIndex = startRun;
            int charIndex = startChar;

            while (remaining > 0 &&
                   runIndex < texts.Count)
            {
                Text? text = texts[runIndex];

                if (text == null)
                {
                    runIndex++;
                    charIndex = 0;
                    continue;
                }

                string value = text.Text;

                int available =
                    value.Length - charIndex;

                int removeCount =
                    Math.Min(remaining, available);

                string before =
                    value.Substring(0, charIndex);

                string after =
                    value.Substring(
                        charIndex + removeCount);

                text.Text = before + after;

                remaining -= removeCount;

                runIndex++;
                charIndex = 0;
            }

            // Chèn tên vào Run chứa ký tự đầu tiên
            Text? targetText =
                texts[startRun];

            if (targetText != null)
            {
                string current = targetText.Text;

                int insertPosition =
                    startChar;

                targetText.Text =
                    current.Insert(
                        insertPosition,
                        names[nameIndex]);
            }
        }
    }
}