using System;

namespace ExcelImportApi.Excel.Attributes;

/// <summary>
/// Marks a property as required (non-null / non-empty).
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelRequiredAttribute : Attribute
{
    public string? ErrorMessage { get; set; }
}
