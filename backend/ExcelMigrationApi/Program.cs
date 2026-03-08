using System.Text.Encodings.Web;
using System.Text.Json;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using ExcelMigrationApi.Services;
using OfficeOpenXml;

// Load .env file if exists (check multiple possible locations)
foreach (var envPath in new[]
{
    Path.Combine(Directory.GetCurrentDirectory(), ".env"),
    Path.Combine(Directory.GetCurrentDirectory(), "..", ".env"),
    Path.Combine(Directory.GetCurrentDirectory(), "..", "..", ".env"),
})
{
    if (File.Exists(envPath))
    {
        foreach (var line in File.ReadAllLines(envPath))
        {
            var parts = line.Split('=', 2);
            if (parts.Length == 2 && !string.IsNullOrWhiteSpace(parts[0]) && !parts[0].TrimStart().StartsWith("#"))
                Environment.SetEnvironmentVariable(parts[0].Trim(), parts[1].Trim());
        }
        break;
    }
}

// Configure EPPlus license (NonCommercial)
ExcelPackage.License.SetNonCommercialOrganization("ExcelMigrationOS");

var builder = WebApplication.CreateBuilder(args);

// Add CORS for Next.js dev server
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.WithOrigins("http://localhost:3000")
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

// Add Clerk JWT authentication
var clerkDomain = Environment.GetEnvironmentVariable("CLERK_DOMAIN") ?? "";
builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options =>
    {
        options.Authority = $"https://{clerkDomain}";
        options.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = true,
            ValidIssuer = $"https://{clerkDomain}",
            ValidateAudience = false,
            ValidateLifetime = true,
            NameClaimType = "sub",
        };
    });
builder.Services.AddAuthorization();

// Add controllers with camelCase JSON
builder.Services.AddControllers()
    .AddJsonOptions(options =>
    {
        options.JsonSerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
        options.JsonSerializerOptions.Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping;
    });

// Register services
builder.Services.AddSingleton<ScanService>();
builder.Services.AddSingleton<ExtractService>();
builder.Services.AddSingleton<ConvertService>();
builder.Services.AddSingleton<UploadService>();
builder.Services.AddSingleton<DeployService>();

builder.Services.AddOpenApi();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

app.UseCors();
app.UseAuthentication();
app.UseAuthorization();

app.MapControllers();

app.Run();
