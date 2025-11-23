using System.IO;
using ClosedXML.Excel;
using ExcelImportApi.GraphQL;

namespace ExcelImportApi.Services;

public class TemplateService
{
    public TemplateDownloadPayload GenerateCountriesTemplate()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Countries");

        // Header row
        ws.Cell(1, 1).Value = "Code";
        ws.Cell(1, 2).Value = "Name";
        ws.Cell(1, 3).Value = "IsActive";

        // Example row (optional)
        ws.Cell(2, 1).Value = "SG";
        ws.Cell(2, 2).Value = "Singapore";
        ws.Cell(2, 3).Value = "TRUE";

        ws.Columns().AdjustToContents();

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        var bytes = ms.ToArray();

        return new TemplateDownloadPayload
        {
            FileName = "countries_template.xlsx",
            ContentBase64 = Convert.ToBase64String(bytes)
        };
    }

    public (byte[] Bytes, string FileName, string ContentType) GenerateCountriesTemplateFile()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Countries");

        // Header row (titles)
        ws.Cell(1, 1).Value = "Code";
        ws.Cell(1, 2).Value = "Name";
        ws.Cell(1, 3).Value = "IsActive";
        ws.Cell(1, 4).Value = "StartDate";


        // Add header instruction messages
        // Col A: Code
        ws.Cell(1, 1).GetComment().AddText(
            "Code must be numeric only.\n" +
            "Example: 1001\n" +
            "No letters or special characters allowed."
        );

        // Col B: Name
        ws.Cell(1, 2).GetComment().AddText(
            "Name can be any text.\n" +
            "Example: Singapore"
        );

        // Col C: IsActive
        ws.Cell(1, 3).GetComment().AddText(
            "Allowed values:\nTRUE, FALSE, YES, NO, 1, 0"
        );

        // Col D: StartDate
        ws.Cell(1, 4).GetComment().AddText(
            "Must be a valid date.\n" +
            "Format: DDMMYYYY\n" +
            "Example: 01012025"
        );

        ws.Cell(2, 1).Value = 1001;                 // numeric
        ws.Cell(2, 2).Value = "Singapore";          // name
        ws.Cell(2, 3).Value = "TRUE";               // boolean
        ws.Cell(2, 4).Value = new DateTime(2025, 1, 1); // date

        // ==========================
        // Format date column
        // ==========================
        var dateColumn = ws.Column(4);
        dateColumn.Style.DateFormat.Format = "dd-MM-yyyy";

        // Data validations
        // CODE → numeric only
        var codeRange = ws.Range("A2:A1048576");
        var codeValidation = codeRange.CreateDataValidation();
        codeValidation.WholeNumber.Between(0, int.MaxValue);
        codeValidation.IgnoreBlanks = true;
        codeValidation.ErrorTitle = "Invalid Code";
        codeValidation.ErrorMessage = "Code must be numeric only.";
        codeValidation.ShowErrorMessage = true;


        // DATE → valid date (DDMMYYYY format applied separately)
        var dateRange = ws.Range("D2:D1048576");
        var dateValidation = dateRange.CreateDataValidation();
        dateValidation.Date.Between(new DateTime(2000, 1, 1), new DateTime(2100, 12, 31));
        dateValidation.IgnoreBlanks = true;
        dateValidation.ErrorTitle = "Invalid Date";
        dateValidation.ErrorMessage = "Please enter a valid date in DDMMYYYY format.";
        dateValidation.ShowErrorMessage = true;


        // Autofit columns + ensure .GetComment()s visible
        ws.Columns().AdjustToContents();
        ws.Column(1).Width += 2; // optional: give space for .GetComment() indicator

        // Generate file     
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        var bytes = ms.ToArray();

        const string contentType =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        return (bytes, "countries_template.xlsx", contentType);
    }


}
