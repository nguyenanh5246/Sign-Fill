using ClosedXML.Excel;
using SignFill.Models;
using System.IO;

namespace SignFill.Services;

public class ExcelReader
{
    public List<DataRow> Read(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException(
                "Không tìm thấy file Excel.",
                filePath);
        }

        using XLWorkbook workbook =
            new(filePath);

        if (!workbook.Worksheets.Any())
        {
            throw new InvalidDataException(
                "File Excel không có worksheet.");
        }

        var worksheet =
            workbook.Worksheets.First();

        var usedRange =
            worksheet.RangeUsed();

        if (usedRange == null)
        {
            throw new InvalidDataException(
                "Worksheet Excel đang trống.");
        }

        int firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
        int lastRow = usedRange.RangeAddress.LastAddress.RowNumber;

        int firstColumn =
            usedRange.RangeAddress.FirstAddress.ColumnNumber;

        int lastColumn =
            usedRange.RangeAddress.LastAddress.ColumnNumber;

        // Đọc header
        List<string> headers = new();

        for (int column = firstColumn;
             column <= lastColumn;
             column++)
        {
            string header =
                worksheet.Cell(firstRow, column)
                         .GetString()
                         .Trim();

            if (string.IsNullOrWhiteSpace(header))
            {
                throw new InvalidDataException(
                    $"Cột {column} không có tên header.");
            }

            if (headers.Contains(
                    header,
                    StringComparer.Ordinal))
            {
                throw new InvalidDataException(
                    $"Header '{header}' bị trùng.");
            }

            headers.Add(header);
        }

        // Đọc từng row
        List<DataRow> rows = new();

        for (int row = firstRow + 1;
             row <= lastRow;
             row++)
        {
            DataRow dataRow = new();

            bool hasData = false;

            for (int column = firstColumn;
                 column <= lastColumn;
                 column++)
            {
                string value =
                    worksheet.Cell(row, column)
                             .GetFormattedString()
                             .Trim();

                if (!string.IsNullOrWhiteSpace(value))
                {
                    hasData = true;
                }

                string header =
                    headers[column - firstColumn];

                dataRow.Set(header, value);
            }

            // Bỏ qua dòng hoàn toàn trống
            if (hasData)
            {
                rows.Add(dataRow);
            }
        }

        if (rows.Count == 0)
        {
            throw new InvalidDataException(
                "Excel không có dữ liệu người dùng.");
        }

        return rows;
    }
}