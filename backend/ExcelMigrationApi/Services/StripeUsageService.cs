using System.Net.Http.Headers;
using System.Text.Json;

namespace ExcelMigrationApi.Services;

/// <summary>
/// Reports token usage to Stripe Billing Meter Events API (server-side, guaranteed delivery).
/// </summary>
public class StripeUsageService
{
    private readonly HttpClient _httpClient;
    private readonly string _secretKey;
    private readonly string _meterEventName;
    private readonly string _meterId;

    private readonly ILogger<StripeUsageService> _logger;

    public StripeUsageService(ILogger<StripeUsageService> logger)
    {
        _logger = logger;
        _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
        _secretKey = Environment.GetEnvironmentVariable("STRIPE_SECRET_KEY") ?? string.Empty;
        _meterEventName = Environment.GetEnvironmentVariable("STRIPE_METER_EVENT_NAME") ?? "ai_tokens";
        _meterId = Environment.GetEnvironmentVariable("STRIPE_METER_ID") ?? string.Empty;
    }

    /// <summary>
    /// Report token usage to Stripe. Throws on failure so callers can block the response.
    /// Uses idempotency key to prevent double-billing on retries.
    /// </summary>
    public async Task ReportUsage(string stripeCustomerId, int totalTokens, string? idempotencyKey = null)
    {
        if (string.IsNullOrEmpty(_secretKey))
            throw new InvalidOperationException("STRIPE_SECRET_KEY is not configured");

        if (string.IsNullOrEmpty(stripeCustomerId) || totalTokens <= 0)
            return;

        var request = new HttpRequestMessage(HttpMethod.Post, "https://api.stripe.com/v1/billing/meter_events");
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _secretKey);

        // Idempotency key prevents double-billing if this request is retried
        if (!string.IsNullOrEmpty(idempotencyKey))
        {
            request.Headers.Add("Idempotency-Key", idempotencyKey);
        }

        var timestamp = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
        request.Content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["event_name"] = _meterEventName,
            ["payload[value]"] = totalTokens.ToString(),
            ["payload[stripe_customer_id]"] = stripeCustomerId,
            ["timestamp"] = timestamp.ToString(),
        });

        var response = await _httpClient.SendAsync(request);
        if (!response.IsSuccessStatusCode)
        {
            _logger.LogError("Stripe usage report failed with status {StatusCode} for customer {CustomerId}", response.StatusCode, stripeCustomerId);
            throw new Exception($"Stripe usage report failed ({response.StatusCode})");
        }
    }

    /// <summary>
    /// Get total token usage for the current billing period (month) from Stripe Billing Meter.
    /// Throws on failure so callers can fail-closed (block request rather than allow unbilled usage).
    /// </summary>
    public async Task<int> GetMonthlyUsage(string stripeCustomerId)
    {
        if (string.IsNullOrEmpty(_secretKey))
            throw new InvalidOperationException("STRIPE_SECRET_KEY is not configured");

        if (string.IsNullOrEmpty(_meterId))
            throw new InvalidOperationException("STRIPE_METER_ID is not configured");

        if (string.IsNullOrEmpty(stripeCustomerId))
            throw new ArgumentException("stripeCustomerId is required");

        // Start of current month (UTC)
        var now = DateTime.UtcNow;
        var startOfMonth = new DateTime(now.Year, now.Month, 1, 0, 0, 0, DateTimeKind.Utc);
        var startTimestamp = new DateTimeOffset(startOfMonth).ToUnixTimeSeconds();
        var endTimestamp = new DateTimeOffset(now).ToUnixTimeSeconds();

        var url = $"https://api.stripe.com/v1/billing/meters/{_meterId}/event_summaries" +
                  $"?customer={stripeCustomerId}" +
                  $"&start_time={startTimestamp}" +
                  $"&end_time={endTimestamp}";

        var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _secretKey);

        var response = await _httpClient.SendAsync(request);
        if (!response.IsSuccessStatusCode)
        {
            _logger.LogError("Stripe usage query failed with status {StatusCode} for customer {CustomerId}", response.StatusCode, stripeCustomerId);
            throw new Exception($"Stripe usage query failed ({response.StatusCode})");
        }

        var body = await response.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(body);
        var root = doc.RootElement;

        if (!root.TryGetProperty("data", out var dataArr) || dataArr.GetArrayLength() == 0)
            return 0;

        int total = 0;
        foreach (var summary in dataArr.EnumerateArray())
        {
            if (summary.TryGetProperty("aggregated_value", out var val))
            {
                total += (int)val.GetDouble();
            }
        }

        return total;
    }
}
