using System.Text.Encodings.Web;
using System.Text.Json;
using System.Threading.RateLimiting;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.RateLimiting;
using Microsoft.IdentityModel.Tokens;
using ExcelMigrationApi.Services;

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

// Fail-fast: verify required environment variables at startup
var requiredEnvVars = new[] { "CLERK_DOMAIN", "CLERK_SECRET_KEY", "ANTHROPIC_API_KEY", "STRIPE_SECRET_KEY", "STRIPE_METER_EVENT_NAME", "STRIPE_METER_ID" };
var missingVars = requiredEnvVars.Where(v => string.IsNullOrEmpty(Environment.GetEnvironmentVariable(v))).ToList();
if (missingVars.Count > 0)
{
    Console.Error.WriteLine($"FATAL: Missing required environment variables: {string.Join(", ", missingVars)}");
    Console.Error.WriteLine("Set these in .env or system environment before starting the server.");
    Environment.Exit(1);
}

var builder = WebApplication.CreateBuilder(args);

// CORS: allow configured origins (production + dev)
var allowedOrigins = new List<string> { "http://localhost:3000" };
var productionOrigin = Environment.GetEnvironmentVariable("FRONTEND_ORIGIN");
if (!string.IsNullOrEmpty(productionOrigin))
{
    allowedOrigins.Add(productionOrigin);
}

builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.WithOrigins(allowedOrigins.ToArray())
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
            ValidateAudience = true,
            ValidAudience = clerkDomain,
            ValidateLifetime = true,
            NameClaimType = "sub",
        };
    });
builder.Services.AddAuthorization();

// Rate limiting
builder.Services.AddRateLimiter(options =>
{
    options.RejectionStatusCode = 429;

    // Global per-user policy: 30 requests per minute
    options.AddPolicy("per-user", context =>
    {
        var userId = context.User?.FindFirst("sub")?.Value ?? context.Connection.RemoteIpAddress?.ToString() ?? "anonymous";
        return RateLimitPartition.GetTokenBucketLimiter(userId, _ => new TokenBucketRateLimiterOptions
        {
            TokenLimit = 30,
            ReplenishmentPeriod = TimeSpan.FromMinutes(1),
            TokensPerPeriod = 30,
            QueueLimit = 0,
        });
    });

    // Strict policy for expensive AI endpoints: 10 requests per minute
    options.AddPolicy("convert", context =>
    {
        var userId = context.User?.FindFirst("sub")?.Value ?? context.Connection.RemoteIpAddress?.ToString() ?? "anonymous";
        return RateLimitPartition.GetTokenBucketLimiter($"convert-{userId}", _ => new TokenBucketRateLimiterOptions
        {
            TokenLimit = 10,
            ReplenishmentPeriod = TimeSpan.FromMinutes(1),
            TokensPerPeriod = 10,
            QueueLimit = 0,
        });
    });
});

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
builder.Services.AddSingleton<ClerkService>();
builder.Services.AddSingleton<StripeUsageService>();

builder.Services.AddOpenApi();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

if (!app.Environment.IsDevelopment())
{
    app.UseHttpsRedirection();
    app.UseHsts();
}

// Global exception handler — prevents stack trace leakage in production
app.UseExceptionHandler(errorApp =>
{
    errorApp.Run(async context =>
    {
        context.Response.StatusCode = 500;
        context.Response.ContentType = "application/json";
        var logger = context.RequestServices.GetRequiredService<ILoggerFactory>().CreateLogger("GlobalExceptionHandler");
        var exFeature = context.Features.Get<Microsoft.AspNetCore.Diagnostics.IExceptionHandlerFeature>();
        if (exFeature?.Error != null)
        {
            logger.LogError(exFeature.Error, "Unhandled exception on {Method} {Path}", context.Request.Method, context.Request.Path);
        }
        await context.Response.WriteAsJsonAsync(new { error = "サーバー内部エラーが発生しました" });
    });
});

// Security headers
app.Use(async (context, next) =>
{
    context.Response.Headers["X-Content-Type-Options"] = "nosniff";
    context.Response.Headers["X-Frame-Options"] = "DENY";
    context.Response.Headers["Referrer-Policy"] = "strict-origin-when-cross-origin";
    if (!app.Environment.IsDevelopment())
    {
        context.Response.Headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains";
    }
    await next();
});

app.UseCors();
app.UseAuthentication();
app.UseAuthorization();
app.UseRateLimiter();

app.MapControllers();

app.Run();
