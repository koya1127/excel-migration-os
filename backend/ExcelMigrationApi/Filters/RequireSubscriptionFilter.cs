using ExcelMigrationApi.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.Extensions.Logging;

namespace ExcelMigrationApi.Filters;

/// <summary>
/// Action filter that verifies the user has an active Stripe subscription
/// and has not exceeded the monthly token cap before allowing access to billable endpoints.
/// </summary>
public class RequireSubscriptionAttribute : TypeFilterAttribute
{
    public RequireSubscriptionAttribute() : base(typeof(RequireSubscriptionFilter)) { }
}

public class RequireSubscriptionFilter : IAsyncActionFilter
{
    private readonly ClerkService _clerkService;
    private readonly StripeUsageService _stripeUsageService;
    private readonly ILogger<RequireSubscriptionFilter> _logger;

    // Monthly cap: 10M tokens (~¥30,000). Configurable via env var.
    private static readonly int MonthlyTokenCap = int.TryParse(
        Environment.GetEnvironmentVariable("MONTHLY_TOKEN_CAP"), out var cap) ? cap : 10_000_000;

    public RequireSubscriptionFilter(ClerkService clerkService, StripeUsageService stripeUsageService, ILogger<RequireSubscriptionFilter> logger)
    {
        _clerkService = clerkService;
        _stripeUsageService = stripeUsageService;
        _logger = logger;
    }

    public async Task OnActionExecutionAsync(ActionExecutingContext context, ActionExecutionDelegate next)
    {
        var userId = context.HttpContext.User.FindFirst("sub")?.Value
                  ?? context.HttpContext.User.FindFirst(System.Security.Claims.ClaimTypes.NameIdentifier)?.Value;

        if (string.IsNullOrEmpty(userId))
        {
            context.Result = new UnauthorizedObjectResult(new { error = "認証情報が取得できません" });
            return;
        }

        var meta = await _clerkService.GetUserMeta(userId);

        if (meta == null || meta.SubscriptionStatus != "active")
        {
            context.Result = new ObjectResult(new { error = "有効なサブスクリプションが必要です。料金プランページからご契約ください。" })
            {
                StatusCode = 403
            };
            return;
        }

        // Check monthly token usage cap
        if (MonthlyTokenCap > 0 && !string.IsNullOrEmpty(meta.StripeCustomerId))
        {
            try
            {
                var usedTokens = await _stripeUsageService.GetMonthlyUsage(meta.StripeCustomerId);
                if (usedTokens >= MonthlyTokenCap)
                {
                    context.Result = new ObjectResult(new
                    {
                        error = $"今月のトークン使用量が上限（{MonthlyTokenCap:N0}トークン）に達しました。来月までお待ちください。"
                    })
                    {
                        StatusCode = 429
                    };
                    return;
                }
            }
            catch (Exception ex)
            {
                // Fail-closed: block request if usage check fails to prevent unbilled usage
                _logger.LogError(ex, "Usage check failed for user {UserId}", userId);
                context.Result = new ObjectResult(new { error = "使用量の確認に失敗しました。しばらくしてから再度お試しください。" })
                {
                    StatusCode = 503
                };
                return;
            }
        }

        // Store metadata in HttpContext for downstream use (usage reporting)
        context.HttpContext.Items["ClerkUserMeta"] = meta;
        context.HttpContext.Items["ClerkUserId"] = userId;

        await next();
    }
}
