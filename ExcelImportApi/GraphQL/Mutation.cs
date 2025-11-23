using ExcelImportApi.Models;
using ExcelImportApi.Services;
using HotChocolate;
using HotChocolate.Types;

namespace ExcelImportApi.GraphQL;

public class ImportJobPayload
{
    public ImportJobPayload(ImportJob job)
    {
        Job = job;
    }

    public ImportJob Job { get; }
}

public class Mutation
{
    // 3) Upload Excel
    public async Task<ImportJobPayload> UploadCountriesExcelAsync(
        [GraphQLType(typeof(NonNullType<UploadType>))] IFile file,
        [Service] ImportService importService,
        CancellationToken cancellationToken)
    {
        // In real app, get from auth
        var currentUser = "system-user";

        using var stream = file.OpenReadStream();
        var job = await importService.ImportCountriesAsync(
            file.Name,
            stream,
            currentUser,
            cancellationToken);

        return new ImportJobPayload(job);
    }
}
