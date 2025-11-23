namespace ExcelImportApi.GraphQL;

public class TemplateDownloadPayload
{
    public string FileName { get; set; } = default!;
    public string ContentType { get; set; } =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    // Base64 encoded .xlsx file
    public string ContentBase64 { get; set; } = default!;
}
