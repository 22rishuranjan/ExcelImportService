using System;

namespace ExcelImportApi.Excel.Attributes;

/// <summary>
/// Maps a model property to a specific Excel column.
/// You can use ColumnIndex (1-based) and/or HeaderName.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelColumnAttribute : Attribute
{
    public ExcelColumnAttribute(string headerName)
    {
        HeaderName = headerName;
    }

    /// <summary>
    /// Human-readable column name (header text).
    /// </summary>
    public string HeaderName { get; }

    /// <summary>
    /// 1-based column index in the sheet (A=1, B=2, ...).
    /// </summary>
    public int ColumnIndex { get; set; }
}
