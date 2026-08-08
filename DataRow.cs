namespace SignFill.Models;

public class DataRow
{
    private readonly Dictionary<string, string> _fields =
        new(StringComparer.Ordinal);

    public IReadOnlyDictionary<string, string> Fields => _fields;

    public void Set(string fieldName, string value)
    {
        if (string.IsNullOrWhiteSpace(fieldName))
            throw new ArgumentException(
                "Tên field không được để trống.",
                nameof(fieldName));

        _fields[fieldName] = value ?? string.Empty;
    }

    public bool Contains(string fieldName)
    {
        return _fields.ContainsKey(fieldName);
    }

    public bool TryGetValue(
        string fieldName,
        out string value)
    {
        return _fields.TryGetValue(
            fieldName,
            out value!);
    }

    public string Get(string fieldName)
    {
        if (!_fields.TryGetValue(
                fieldName,
                out string? value))
        {
            throw new KeyNotFoundException(
                $"Không tìm thấy field '{fieldName}'.");
        }

        return value;
    }
}