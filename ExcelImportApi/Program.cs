using ExcelImportApi.Data;
using ExcelImportApi.GraphQL;
using ExcelImportApi.Services;
using HotChocolate.AspNetCore.Voyager;
using HotChocolate.Types;
using MongoDB.Driver;

var builder = WebApplication.CreateBuilder(args);

// Mongo config
builder.Services.Configure<MongoSettings>(
    builder.Configuration.GetSection("Mongo"));

builder.Services.AddSingleton<MongoDbContext>();
builder.Services.AddScoped<ImportService>();
builder.Services.AddScoped<TemplateService>();
builder.Services.AddScoped<ErrorReportService>();

// GraphQL
builder.Services
    .AddGraphQLServer()
    .AddQueryType<Query>()
    .AddMutationType<Mutation>()
    .AddType<UploadType>()
    .ModifyRequestOptions(opt =>
    {
        opt.IncludeExceptionDetails = true;
    });

var app = builder.Build();

// REST endpoints

// Download Excel template
app.MapGet("/api/countries/template", (TemplateService templateService) =>
{
    var (bytes, fileName, contentType) = templateService.GenerateCountriesTemplateFile();

    return Results.File(
        fileContents: bytes,
        contentType: contentType,
        fileDownloadName: fileName
    );
});

app.MapPost("/api/countries/upload", async (
    HttpRequest request,
    ImportService importService,
    CancellationToken cancellationToken) =>
{
    if (!request.HasFormContentType)
        return Results.BadRequest("Content-Type must be multipart/form-data.");

    var form = await request.ReadFormAsync(cancellationToken);
    var file = form.Files["file"];

    if (file is null || file.Length == 0)
        return Results.BadRequest("File 'file' is required and cannot be empty.");

    var currentUser = "system-user";

    await using var stream = file.OpenReadStream();
    var job = await importService.ImportCountriesAsync(
        file.FileName,
        stream,
        currentUser,
        cancellationToken);

    return Results.Ok(job);
});


// Download the original Excel with a Status column based on errors
app.MapGet("/api/countries/import/{id}", async (
    string id,
    IServiceProvider sp,
    CancellationToken cancellationToken) =>
{
    // Resolve services manually from DI
    var db = sp.GetRequiredService<MongoDbContext>();
    var errorReportService = sp.GetRequiredService<ErrorReportService>();

    var job = await db.ImportJobs
        .Find(j => j.Id == id)
        .FirstOrDefaultAsync(cancellationToken);

    if (job is null)
    {
        return Results.NotFound($"Import job '{id}' not found.");
    }

    if (job.Errors == null || job.Errors.Count == 0)
    {
        return Results.BadRequest("This import job has no recorded errors.");
    }

    var (bytes, fileName, contentType) = errorReportService.GenerateAnnotatedExcel(job);

    return Results.File(bytes, contentType, fileName);
});



// GraphQL endpoint
app.MapGraphQL("/graphql");

// Optional Voyager
app.UseVoyager("/graphql", "/voyager");

app.Run();
