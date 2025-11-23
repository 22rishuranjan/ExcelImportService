using ExcelImportApi.Models;
using Microsoft.Extensions.Options;
using MongoDB.Driver;

namespace ExcelImportApi.Data;

public class MongoSettings
{
    public string ConnectionString { get; set; } = default!;
    public string DatabaseName { get; set; } = default!;
}

public class MongoDbContext
{
    private readonly IMongoDatabase _db;

    public MongoDbContext(IOptions<MongoSettings> options)
    {
        var client = new MongoClient(options.Value.ConnectionString);
        _db = client.GetDatabase(options.Value.DatabaseName);
    }

    public IMongoCollection<Country> Countries => _db.GetCollection<Country>("countries");
    public IMongoCollection<ImportJob> ImportJobs => _db.GetCollection<ImportJob>("importJobs");
}
