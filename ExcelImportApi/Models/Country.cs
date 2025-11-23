using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;

namespace ExcelImportApi.Models;

public class Country
{
    [BsonId]
    [BsonRepresentation(BsonType.ObjectId)]
    public string Id { get; set; } = default!;

    public string Code { get; set; } = default!;
    public string Name { get; set; } = default!;

    public bool IsActive { get; set; }

    public DateTime CreatedAtUtc { get; set; }
    public string CreatedBy { get; set; } = default!;
}
