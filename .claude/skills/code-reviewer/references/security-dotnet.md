# Security Best Practices for C# (.NET)

## 1. Authentication & Authorization
- **JWT Validation**: Ensure `IssuerSigningKey`, `ValidateIssuer`, `ValidateAudience`, and `ValidateLifetime` are explicitly set and not disabled.
- **Middleware Order**: Always ensure `UseAuthentication()` is called BEFORE `UseAuthorization()`.
- **Policy-based Auth**: Check if controllers use `[Authorize(Policy = "...")]` rather than hardcoding role checks.

## 2. Data Access Security
- **Entity Framework Core**: Use parameterized queries (default in LINQ). Flag any raw SQL via `FromSqlRaw` if user input is concatenated.
- **Sensitive Data**: Ensure passwords or secrets are never stored in plain text. Use `Identity` or bcrypt.
- **Output Encoding**: Use `HtmlEncoder` or `JavaScriptEncoder` for manual output (though Razor/Blazor does this by default).

## 3. Configuration & Secrets
- **No Secrets in Source**: Flag any `.json` files or `.cs` constants containing API keys or connection strings.
- **Environment Variables**: Verify that secrets are pulled from `Environment` or `Secret Manager` (locally) or `Key Vault` (production).

## 4. API Security
- **CORS Policy**: Ensure CORS is not set to `AllowAnyOrigin` in production.
- **Rate Limiting**: Check if any rate-limiting middleware is used for public APIs.
- **Input Validation**: Use `[Required]`, `[StringLength]`, etc., on DTOs to prevent over-posting or buffer-related issues.
