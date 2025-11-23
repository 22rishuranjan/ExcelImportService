using ClosedXML.Excel;
using ExcelImportApi.Data;
using ExcelImportApi.Excel.Validation;
using ExcelImportApi.Models;
using MongoDB.Driver;

namespace ExcelImportApi.Services;

public class ImportService
{
    private readonly MongoDbContext _db;

    public ImportService(MongoDbContext db)
    {
        _db = db;
    }

    public async Task<ImportJob> ImportCountriesAsync(
        string fileName,
        Stream fileStream,
        string currentUser,
        CancellationToken cancellationToken = default)
    {
        // Read the incoming stream into a byte[] so we can:
        // 1) Store it on the job (OriginalFile)
        // 2) Use it multiple times (ClosedXML)
        byte[] fileBytes;
        using (var ms = new MemoryStream())
        {
            await fileStream.CopyToAsync(ms, cancellationToken);
            fileBytes = ms.ToArray();
        }

        var job = new ImportJob
        {
            FileName = fileName,
            Status = ImportStatus.Running,
            StartedAtUtc = DateTime.UtcNow,
            OriginalFile = fileBytes  // 👈 store entire Excel
        };

        await _db.ImportJobs.InsertOneAsync(job, cancellationToken: cancellationToken);

        var errors = new List<ImportError>();
        var countriesToInsert = new List<Country>();

        try
        {
            using var workbookStream = new MemoryStream(fileBytes);
            using var workbook = new XLWorkbook(workbookStream);
            var ws = workbook.Worksheets.First();

            var firstDataRow = 2;
            var lastRow = ws.LastRowUsed().RowNumber();
            job.TotalRows = lastRow - firstDataRow + 1;

            for (int rowNum = firstDataRow; rowNum <= lastRow; rowNum++)
            {
                var row = ws.Row(rowNum);

                var validationResult =
                    ExcelRowValidator.ValidateRow<CountryUploadRow>(row, rowNum);

                if (validationResult.Errors.Any())
                {
                    errors.AddRange(validationResult.Errors);
                    continue;
                }

                var uploadRow = validationResult.Model!;

                var country = new Country
                {
                    Code = uploadRow.Code.ToString(),
                    Name = uploadRow.Name,
                    IsActive = uploadRow.IsActive,
                    CreatedAtUtc = DateTime.UtcNow,
                    CreatedBy = currentUser
                    // StartDate is available on uploadRow.StartDate if you want to store it
                };

                countriesToInsert.Add(country);
            }

            if (errors.Any())
            {
                job.Status = ImportStatus.Failed;
                job.Errors = errors;
                job.FailureCount = errors
                    .Select(e => e.Row)
                    .Distinct()
                    .Count();
                job.SuccessCount = countriesToInsert.Count;
            }
            else
            {
                if (countriesToInsert.Any())
                {
                    await _db.Countries.InsertManyAsync(
                        countriesToInsert,
                        cancellationToken: cancellationToken);
                }

                job.Status = ImportStatus.Completed;
                job.SuccessCount = countriesToInsert.Count;
                job.FailureCount = 0;
            }

            job.CompletedAtUtc = DateTime.UtcNow;

            var filter = Builders<ImportJob>.Filter.Eq(j => j.Id, job.Id);
            await _db.ImportJobs.ReplaceOneAsync(filter, job, cancellationToken: cancellationToken);

            return job;
        }
        catch (Exception ex)
        {
            job.Status = ImportStatus.Failed;
            job.Errors.Add(new ImportError
            {
                Row = 0,
                Column = "",
                Message = $"Unexpected error: {ex.Message}"
            });
            job.CompletedAtUtc = DateTime.UtcNow;

            var filter = Builders<ImportJob>.Filter.Eq(j => j.Id, job.Id);
            await _db.ImportJobs.ReplaceOneAsync(filter, job, cancellationToken: cancellationToken);

            return job;
        }
    }
}
