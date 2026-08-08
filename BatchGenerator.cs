using SignFill.Models;
using System.IO;

namespace SignFill.Services;

public class BatchGenerator
{
    private readonly WordGenerator _wordGenerator;
    private readonly FileNameGenerator _fileNameGenerator;

    public BatchGenerator()
    {
        _wordGenerator = new WordGenerator();
        _fileNameGenerator = new FileNameGenerator();
    }

    public int Generate(
        string templatePath,
        string outputDirectory,
        List<DataRow> rows,
        string fileNamePattern)
    {
        if (rows == null || rows.Count == 0)
        {
            throw new InvalidDataException(
                "Không có dữ liệu để generate.");
        }

        if (string.IsNullOrWhiteSpace(fileNamePattern))
        {
            throw new ArgumentException(
                "Mẫu tên file không được để trống.",
                nameof(fileNamePattern));
        }

        Directory.CreateDirectory(outputDirectory);

        // -------------------------------------------------
        // PHASE 1:
        // Tạo toàn bộ tên file trước khi generate.
        // -------------------------------------------------

        List<string> fileNames = new();

        for (int i = 0; i < rows.Count; i++)
        {
            string fileName =
                _fileNameGenerator.Generate(
                    fileNamePattern,
                    rows[i]);

            fileNames.Add(fileName + ".docx");
        }

        // -------------------------------------------------
        // PHASE 2:
        // Kiểm tra trùng tên file.
        // -------------------------------------------------

        var duplicateGroups =
            fileNames
                .GroupBy(
                    name => name,
                    StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1)
                .ToList();

        if (duplicateGroups.Count > 0)
        {
            List<string> duplicateNames =
                duplicateGroups
                    .Select(group => group.Key)
                    .ToList();

            string message =
                "Phát hiện tên file bị trùng:\n\n" +
                string.Join(
                    "\n",
                    duplicateNames) +
                "\n\n" +
                "Vui lòng thêm một trường khác " +
                "vào mẫu tên file để đảm bảo " +
                "mỗi hợp đồng có tên duy nhất.";

            throw new InvalidDataException(message);
        }

        // -------------------------------------------------
        // PHASE 3:
        // Generate Word.
        // -------------------------------------------------

        int generatedCount = 0;

        for (int i = 0; i < rows.Count; i++)
        {
            string outputPath =
                Path.Combine(
                    outputDirectory,
                    fileNames[i]);

            _wordGenerator.Generate(
                templatePath,
                outputPath,
                rows[i]);

            generatedCount++;
        }

        return generatedCount;
    }
}