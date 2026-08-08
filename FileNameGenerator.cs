using SignFill.Models;
using System.IO;
using System.Text.RegularExpressions;

namespace SignFill.Services;

public class FileNameGenerator
{
    private static readonly Regex FieldRegex =
        new(@"\{([^{}]+)\}",
            RegexOptions.Compiled);

    public string Generate(
        string pattern,
        DataRow row)
    {
        if (string.IsNullOrWhiteSpace(pattern))
        {
            throw new ArgumentException(
                "Tên file không được để trống.",
                nameof(pattern));
        }

        if (row == null)
        {
            throw new ArgumentNullException(nameof(row));
        }

        string result = pattern;

        MatchCollection matches =
            FieldRegex.Matches(pattern);

        foreach (Match match in matches)
        {
            string fieldName =
                match.Groups[1].Value;

            if (!row.Fields.TryGetValue(
                    fieldName,
                    out string? value))
            {
                throw new InvalidDataException(
                    $"Field '{fieldName}' " +
                    "trong tên file không tồn tại trong Excel.");
            }

            result =
                result.Replace(
                    match.Value,
                    value ?? string.Empty);
        }

        return SanitizeFileName(result);
    }

    private static string SanitizeFileName(
        string fileName)
    {
        foreach (char invalidChar in
                 Path.GetInvalidFileNameChars())
        {
            fileName =
                fileName.Replace(
                    invalidChar,
                    '_');
        }

        return fileName.Trim();
    }
}