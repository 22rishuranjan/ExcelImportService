using ExcelImportApi.Data;
using ExcelImportApi.Models;
using ExcelImportApi.Services;
using MongoDB.Bson;
using MongoDB.Driver;

namespace ExcelImportApi.GraphQL;

public class Query
{
    public async Task<List<Country>> GetCountries([Service] MongoDbContext db)
    {
        return await db.Countries.Find(FilterDefinition<Country>.Empty).ToListAsync();
    }

    // 1) Query result: get import job status/details
    public async Task<ImportJob?> GetImportJob(
        string id,
        [Service] MongoDbContext db)
    {
        return await db.ImportJobs
            .Find(j => j.Id == id)
            .FirstOrDefaultAsync();
    }

    // Optional: list jobs
    public async Task<List<ImportJob>> GetImportJobs([Service] MongoDbContext db)
    {
        return await db.ImportJobs
            .Find(FilterDefinition<ImportJob>.Empty)
            .SortByDescending(j => j.StartedAtUtc)
            .Limit(20)
            .ToListAsync();
    }

    // 2) Download template: returns base64 Excel file
    public TemplateDownloadPayload GetCountriesTemplate(
        [Service] TemplateService templateService)
    {
        return templateService.GenerateCountriesTemplate();
    }
}
