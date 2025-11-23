using System;

namespace ExcelImportApi.Excel.Attributes;

/// <summary>
/// Ensures that the cell can be parsed as a DateTime in the given format, 
/// and optionally enforces min/max year.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelDateAttribute : Attribute
{
    public ExcelDateAttribute(string format)
    {
        Format = format;
    }

    public string Format { get; }
    public int MinYear { get; set; } = 0;
    public int MaxYear { get; set; } = 0;

    public string? ErrorMessage { get; set; }
}
