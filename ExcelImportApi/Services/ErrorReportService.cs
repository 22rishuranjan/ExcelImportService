using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExcelImportApi.Models;

namespace ExcelImportApi.Services;

/// <summary>
/// Creates an Excel file by taking the original upload and appending
/// a 'Status' column with per-row error messages.
/// </summary>
public class ErrorReportService
{
    public (byte[] Bytes, string FileName, string ContentType) GenerateAnnotatedExcel(ImportJob job)
    {
        if (job.OriginalFile == null || job.OriginalFile.Length == 0)
        {
            // Fallback: if original file not stored, just create a simple Errors sheet
            return GenerateSimpleErrorReport(job);
        }

        using var ms = new MemoryStream(job.OriginalFile);
        using var workbook = new XLWorkbook(ms);
        var ws = workbook.Worksheets.First();

        // Determine where to put the Status column (last used column + 1)
        var lastColumn = ws.LastColumnUsed().ColumnNumber();
        var statusColIndex = lastColumn + 1;

        ws.Cell(1, statusColIndex).Value = "Status";
        ws.Cell(1, statusColIndex).Style.Font.Bold = true;

        // Group errors by row
        var errorsByRow = job.Errors
            .GroupBy(e => e.Row)
            .ToDictionary(g => g.Key, g => g.ToList());

        foreach (var kvp in errorsByRow)
        {
            var rowNumber = kvp.Key;
            var rowErrors = kvp.Value;

            var statusText = string.Join(
                "; ",
                rowErrors.Select(e => $"{e.Column}: {e.Message}"));

            var row = ws.Row(rowNumber);
            row.Cell(statusColIndex).Value = statusText;
        }

        ws.Columns().AdjustToContents();

        using var outStream = new MemoryStream();
        workbook.SaveAs(outStream);
        var bytes = outStream.ToArray();

        var safeFileName = $"import_with_status_{job.Id}.xlsx";
        const string contentType =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        return (bytes, safeFileName, contentType);
    }

    private (byte[] Bytes, string FileName, string ContentType) GenerateSimpleErrorReport(ImportJob job)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Errors");

        ws.Cell(1, 1).Value = "Row";
        ws.Cell(1, 2).Value = "Column";
        ws.Cell(1, 3).Value = "Message";

        ws.Row(1).Style.Font.Bold = true;

        var currentRow = 2;
        foreach (var err in job.Errors)
        {
            ws.Cell(currentRow, 1).Value = err.Row;
            ws.Cell(currentRow, 2).Value = err.Column;
            ws.Cell(currentRow, 3).Value = err.Message;
            currentRow++;
        }

        ws.Columns().AdjustToContents();

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        var bytes = ms.ToArray();

        var safeFileName = $"import_errors_{job.Id}.xlsx";
        const string contentType =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        return (bytes, safeFileName, contentType);
    }
}
