using SignFill.Models;
using System.IO;

namespace SignFill.Services;

public class ValidationService
{
    public void Validate(
        List<DataRow> rows,
        List<string> templateFields)
    {
        if (rows == null || rows.Count == 0)
        {
            throw new InvalidDataException(
                "Excel không có dữ liệu.");
        }

        if (templateFields == null ||
            templateFields.Count == 0)
        {
            throw new InvalidDataException(
                "Template không chứa field nào.");
        }

        DataRow firstRow = rows[0];

        List<string> missingFields = new();

        foreach (string templateField in templateFields)
        {
            if (!firstRow.Contains(templateField))
            {
                missingFields.Add(templateField);
            }
        }

        if (missingFields.Count > 0)
        {
            string errorMessage =
                "Các field trong template không tồn tại trong Excel:\n\n" +
                string.Join(
                    "\n",
                    missingFields.Select(
                        field => $"{{{{{field}}}}}"));

            throw new InvalidDataException(
                errorMessage);
        }
    }
}