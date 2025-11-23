using System;

namespace ExcelImportApi.Excel.Attributes;

/// <summary>
/// Ensures that the cell can be parsed as a number.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelNumericAttribute : Attribute
{
    public double? Min { get; set; }
    public double? Max { get; set; }

    public string? ErrorMessage { get; set; }
}
