using System;

namespace ExcelImportApi.Excel.Attributes;

/// <summary>
/// Ensures that the cell value matches one of the allowed values (case-insensitive).
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelAllowedValuesAttribute : Attribute
{
    public ExcelAllowedValuesAttribute(params string[] values)
    {
        Values = values;
    }

    public string[] Values { get; }

    public string? ErrorMessage { get; set; }
}
