using System;

namespace ExcelImportApi.Excel.Attributes;

/// <summary>
/// Ensures that the cell can be parsed as a boolean, using configurable true/false tokens.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelBooleanAttribute : Attribute
{
    public string[] AllowedTrueValues { get; set; } = new[] { "true", "1" };
    public string[] AllowedFalseValues { get; set; } = new[] { "false", "0" };

    public string? ErrorMessage { get; set; }
}
