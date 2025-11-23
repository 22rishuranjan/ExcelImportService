using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace ExcelImportApi.Models;

public enum ImportStatus
{
    Pending,
    Running,
    Completed,
    Failed
}

public class ImportError
{
    public int Row { get; set; }
    public string Column { get; set; } = default!;
    public string Message { get; set; } = default!;
}


public class ImportJob
{
    [BsonId]
    [BsonRepresentation(BsonType.ObjectId)]
    public string Id { get; set; } = default!;

    public string FileName { get; set; } = default!;

    public ImportStatus Status { get; set; }

    public int TotalRows { get; set; }
    public int SuccessCount { get; set; }
    public int FailureCount { get; set; }

    public List<ImportError> Errors { get; set; } = new();

    public DateTime StartedAtUtc { get; set; }
    public DateTime? CompletedAtUtc { get; set; }

    /// <summary>
    /// Original uploaded Excel file bytes.
    /// Used later to regenerate Excel with inline error annotations.
    /// </summary>
    [BsonIgnoreIfNull]
    public byte[]? OriginalFile { get; set; }
}

