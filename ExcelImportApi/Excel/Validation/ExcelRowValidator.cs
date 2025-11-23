using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using ExcelImportApi.Excel.Attributes;
using ExcelImportApi.Models;

namespace ExcelImportApi.Excel.Validation;

/// <summary>
/// Uses reflection + custom attributes to map and validate an Excel row into a model.
/// </summary>
public static class ExcelRowValidator
{
    public static ExcelRowValidationResult<T> ValidateRow<T>(IXLRow row, int rowNumber)
        where T : new()
    {
        var result = new ExcelRowValidationResult<T>();
        var errors = new List<ImportError>();
        var model = new T();

        var type = typeof(T);
        var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);

        foreach (var prop in properties)
        {
            var colAttr = prop.GetCustomAttribute<ExcelColumnAttribute>();
            if (colAttr == null)
            {
                continue; // not mapped to Excel
            }

            var columnIndex = colAttr.ColumnIndex;
            var colName = !string.IsNullOrWhiteSpace(colAttr.HeaderName)
                ? colAttr.HeaderName
                : $"Column {columnIndex}";

            var cell = row.Cell(columnIndex);
            var rawValue = cell.GetString().Trim();

            // REQUIRED
            var requiredAttr = prop.GetCustomAttribute<ExcelRequiredAttribute>();
            if (requiredAttr != null)
            {
                if (string.IsNullOrWhiteSpace(rawValue))
                {
                    errors.Add(new ImportError
                    {
                        Row = rowNumber,
                        Column = colName,
                        Message = requiredAttr.ErrorMessage
                                  ?? $"{colName} is required."
                    });
                    // skip further validation for this property
                    continue;
                }
            }

            // AllowedValues (optional)
            var allowedValuesAttr = prop.GetCustomAttribute<ExcelAllowedValuesAttribute>();
            if (allowedValuesAttr != null && !string.IsNullOrWhiteSpace(rawValue))
            {
                var match = allowedValuesAttr.Values
                    .Any(v => string.Equals(v, rawValue, StringComparison.OrdinalIgnoreCase));

                if (!match)
                {
                    errors.Add(new ImportError
                    {
                        Row = rowNumber,
                        Column = colName,
                        Message = allowedValuesAttr.ErrorMessage
                                  ?? $"{colName} must be one of: {string.Join(", ", allowedValuesAttr.Values)}."
                    });
                    continue;
                }
            }

            // NUMERIC
            var numericAttr = prop.GetCustomAttribute<ExcelNumericAttribute>();
            if (numericAttr != null && !string.IsNullOrWhiteSpace(rawValue))
            {
                if (!double.TryParse(rawValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
                {
                    errors.Add(new ImportError
                    {
                        Row = rowNumber,
                        Column = colName,
                        Message = numericAttr.ErrorMessage
                                  ?? $"{colName} must be a numeric value."
                    });
                    continue;
                }

                if (numericAttr.Min.HasValue && number < numericAttr.Min.Value ||
                    numericAttr.Max.HasValue && number > numericAttr.Max.Value)
                {
                    errors.Add(new ImportError
                    {
                        Row = rowNumber,
                        Column = colName,
                        Message = numericAttr.ErrorMessage
                                  ?? $"{colName} must be between {numericAttr.Min} and {numericAttr.Max}."
                    });
                    continue;
                }

                // Assign numeric to property (int / double / long)
                if (prop.PropertyType == typeof(int))
                    prop.SetValue(model, (int)number);
                else if (prop.PropertyType == typeof(long))
                    prop.SetValue(model, (long)number);
                else if (prop.PropertyType == typeof(double))
                    prop.SetValue(model, number);
                else
                    prop.SetValue(model, number); // fallback
                continue;
            }

            // BOOLEAN
            var boolAttr = prop.GetCustomAttribute<ExcelBooleanAttribute>();
            if (boolAttr != null && !string.IsNullOrWhiteSpace(rawValue))
            {
                var normalized = rawValue.Trim().ToLowerInvariant();

                bool? boolValue = null;

                if (boolAttr.AllowedTrueValues.Any(v => v.Equals(normalized, StringComparison.OrdinalIgnoreCase)))
                    boolValue = true;
                else if (boolAttr.AllowedFalseValues.Any(v => v.Equals(normalized, StringComparison.OrdinalIgnoreCase)))
                    boolValue = false;

                if (!boolValue.HasValue)
                {
                    errors.Add(new ImportError
                    {
                        Row = rowNumber,
                        Column = colName,
                        Message = boolAttr.ErrorMessage
                                  ?? $"{colName} must be a valid boolean value."
                    });
                    continue;
                }

                prop.SetValue(model, boolValue.Value);
                continue;
            }

            // DATE
            var dateAttr = prop.GetCustomAttribute<ExcelDateAttribute>();
            if (dateAttr != null && !string.IsNullOrWhiteSpace(rawValue))
            {
                if (!DateTime.TryParseExact(
                        rawValue,
                        dateAttr.Format,
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out var date))
                {
                    errors.Add(new ImportError
                    {
                        Row = rowNumber,
                        Column = colName,
                        Message = dateAttr.ErrorMessage
                                  ?? $"{colName} must be a valid date in format {dateAttr.Format}."
                    });
                    continue;
                }

                if ((dateAttr.MinYear > 0 && date.Year < dateAttr.MinYear) ||
          (dateAttr.MaxYear > 0 && date.Year > dateAttr.MaxYear))
                {
                    errors.Add(new ImportError
                    {
                        Row = rowNumber,
                        Column = colName,
                        Message = dateAttr.ErrorMessage
                                  ?? $"{colName} must be between years {dateAttr.MinYear} and {dateAttr.MaxYear}."
                    });
                    continue;
                }


                prop.SetValue(model, date);
                continue;
            }

            // If no special validation attributes, just assign raw string to string properties.
            if (prop.PropertyType == typeof(string))
            {
                prop.SetValue(model, rawValue);
            }
        }

        result.Model = errors.Any() ? default : model;
        result.Errors = errors;

        return result;
    }
}
