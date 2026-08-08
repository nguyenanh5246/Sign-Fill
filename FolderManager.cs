using System.IO;

namespace SignFill.Services;

public static class FolderManager
{
    public static string BaseDirectory =>
        AppContext.BaseDirectory;

    public static string InputDirectory =>
        Path.Combine(BaseDirectory, "Input");

    public static string TemplatesDirectory =>
        Path.Combine(BaseDirectory, "Templates");

    public static string OutputDirectory =>
        Path.Combine(BaseDirectory, "Output");

    public static void EnsureDirectories()
    {
        Directory.CreateDirectory(InputDirectory);
        Directory.CreateDirectory(TemplatesDirectory);
        Directory.CreateDirectory(OutputDirectory);
    }
}