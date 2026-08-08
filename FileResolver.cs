using System.IO;

namespace SignFill.Services;

public class FileResolver
{
    public string GetExcelFile()
    {
        return GetSingleFile(
            FolderManager.InputDirectory,
            "*.xlsx",
            "Excel");
    }

    public string GetTemplateFile()
    {
        return GetSingleFile(
            FolderManager.TemplatesDirectory,
            "*.docx",
            "Word template");
    }

    private static string GetSingleFile(
        string directory,
        string searchPattern,
        string fileType)
    {
        if (!Directory.Exists(directory))
        {
            throw new DirectoryNotFoundException(
                $"Không tìm thấy thư mục: {directory}");
        }

        string[] files =
            Directory.GetFiles(
                directory,
                searchPattern);

        if (files.Length == 0)
        {
            throw new InvalidDataException(
                $"Không tìm thấy file {fileType} " +
                $"trong thư mục:\n{directory}");
        }

        if (files.Length > 1)
        {
            throw new InvalidDataException(
                $"Tìm thấy {files.Length} file {fileType} " +
                $"trong thư mục:\n{directory}\n\n" +
                $"Vui lòng chỉ giữ lại đúng 1 file.");
        }

        return files[0];
    }
}