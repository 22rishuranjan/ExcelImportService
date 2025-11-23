using ExcelImportApi.Models;
using System.Collections.Generic;

namespace ExcelImportApi.Excel.Validation;

/// <summary>
/// Result of mapping and validating a single Excel row.
/// </summary>
public class ExcelRowValidationResult<T>
{
    public T? Model { get; set; }

    public List<ImportError> Errors { get; set; } = new();
}
