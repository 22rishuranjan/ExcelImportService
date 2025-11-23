using ExcelImportApi.Excel.Attributes;

namespace ExcelImportApi.Models;

/// <summary>
/// Represents a single row from the Countries Excel upload.
/// Attributes define how to map and validate each column.
/// </summary>
public class CountryUploadRow
{
    [ExcelColumn("Code", ColumnIndex = 1)]
    [ExcelRequired]
    [ExcelNumeric]
    public int Code { get; set; }

    [ExcelColumn("Name", ColumnIndex = 2)]
    [ExcelRequired]
    public string Name { get; set; } = default!;

    [ExcelColumn("IsActive", ColumnIndex = 3)]
    [ExcelRequired]
    [ExcelBoolean(AllowedTrueValues = new[] { "true", "yes", "y", "1" },
                  AllowedFalseValues = new[] { "false", "no", "n", "0" })]
    public bool IsActive { get; set; }

    [ExcelColumn("StartDate", ColumnIndex = 4)]
    [ExcelRequired]
    [ExcelDate("ddMMyyyy", MinYear = 2000, MaxYear = 2100)]
    public DateTime StartDate { get; set; }
}
