using System.Collections.Concurrent;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelMigrationApi.Services;

/// <summary>
/// Calls Clerk Backend API to read user metadata (subscription status, Stripe customer ID, Google tokens).
/// </summary>
public class ClerkService
{
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly string _googleClientId;
    private readonly string _googleClientSecret;

    // Per-user lock to prevent concurrent token refresh races.
    // Entries are cleaned up periodically to prevent unbounded growth.
    private static readonly ConcurrentDictionary<string, (SemaphoreSlim Lock, DateTime LastUsed)> _refreshLocks = new();
    private static readonly Timer _cleanupTimer = new(_ =>
    {
        var cutoff = DateTime.UtcNow.AddMinutes(-30);
        foreach (var kvp in _refreshLocks)
        {
            if (kvp.Value.LastUsed < cutoff && kvp.Value.Lock.CurrentCount == 1)
            {
                _refreshLocks.TryRemove(kvp.Key, out var _removed);
            }
        }
    }, null, TimeSpan.FromMinutes(10), TimeSpan.FromMinutes(10));

    public ClerkService(IHttpClientFactory httpClientFactory)
    {
        _httpClientFactory = httpClientFactory;
        _googleClientId = Environment.GetEnvironmentVariable("GOOGLE_CLIENT_ID") ?? string.Empty;
        _googleClientSecret = Environment.GetEnvironmentVariable("GOOGLE_CLIENT_SECRET") ?? string.Empty;
    }

    public async Task<ClerkUserMeta?> GetUserMeta(string userId)
    {
        if (string.IsNullOrEmpty(userId))
            return null;

        try
        {
            var httpClient = _httpClientFactory.CreateClient("Clerk");
            var res = await httpClient.GetAsync($"users/{userId}");
            if (!res.IsSuccessStatusCode) return null;

            var body = await res.Content.ReadAsStringAsync();
            var user = JsonSerializer.Deserialize<ClerkUserResponse>(body);
            if (user?.PublicMetadata == null) return null;

            return new ClerkUserMeta
            {
                SubscriptionStatus = GetString(user.PublicMetadata, "subscriptionStatus"),
                StripeCustomerId = GetString(user.PublicMetadata, "stripeCustomerId"),
            };
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Get a valid Google access token for the user, refreshing if needed.
    /// Token is fetched from Clerk privateMetadata (never exposed to the browser).
    /// Uses per-user locking to prevent concurrent refresh races.
    /// </summary>
    public async Task<string?> GetGoogleToken(string userId)
    {
        if (string.IsNullOrEmpty(userId))
            return null;

        try
        {
            var httpClient = _httpClientFactory.CreateClient("Clerk");
            var res = await httpClient.GetAsync($"users/{userId}");
            if (!res.IsSuccessStatusCode) return null;

            var body = await res.Content.ReadAsStringAsync();
            var user = JsonSerializer.Deserialize<ClerkUserResponse>(body);
            var privateMeta = user?.PrivateMetadata;
            if (privateMeta == null) return null;

            var accessToken = GetString(privateMeta, "googleAccessToken");
            var refreshToken = GetString(privateMeta, "googleRefreshToken");
            var expiry = GetLong(privateMeta, "googleTokenExpiry");

            if (string.IsNullOrEmpty(accessToken) || string.IsNullOrEmpty(refreshToken))
                return null;

            // If token still valid (5 min buffer), return it
            var nowMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
            if (expiry.HasValue && nowMs < expiry.Value - 5 * 60 * 1000)
                return accessToken;

            // Acquire per-user lock to prevent concurrent refresh
            var entry = _refreshLocks.GetOrAdd(userId, _ => (new SemaphoreSlim(1, 1), DateTime.UtcNow));
            _refreshLocks[userId] = (entry.Lock, DateTime.UtcNow); // Update last-used timestamp
            await entry.Lock.WaitAsync();
            try
            {
                // Re-read metadata in case another request already refreshed
                var httpClient2 = _httpClientFactory.CreateClient("Clerk");
                var res2 = await httpClient2.GetAsync($"users/{userId}");
                if (res2.IsSuccessStatusCode)
                {
                    var body2 = await res2.Content.ReadAsStringAsync();
                    var user2 = JsonSerializer.Deserialize<ClerkUserResponse>(body2);
                    var pm2 = user2?.PrivateMetadata;
                    if (pm2 != null)
                    {
                        var expiry2 = GetLong(pm2, "googleTokenExpiry");
                        var nowMs2 = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
                        if (expiry2.HasValue && nowMs2 < expiry2.Value - 5 * 60 * 1000)
                        {
                            return GetString(pm2, "googleAccessToken");
                        }
                    }
                }

                // Refresh the token
                var refreshed = await RefreshGoogleToken(refreshToken);
                if (refreshed == null) return null;

                // Update Clerk privateMetadata with new token
                await UpdatePrivateMetadata(userId, new Dictionary<string, object>
                {
                    ["googleAccessToken"] = refreshed.Value.AccessToken,
                    ["googleRefreshToken"] = refreshToken,
                    ["googleTokenExpiry"] = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + refreshed.Value.ExpiresIn * 1000,
                    ["googleConnected"] = true,
                });

                return refreshed.Value.AccessToken;
            }
            finally
            {
                entry.Lock.Release();
            }
        }
        catch
        {
            return null;
        }
    }

    private async Task<(string AccessToken, int ExpiresIn)?> RefreshGoogleToken(string refreshToken)
    {
        var content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["client_id"] = _googleClientId,
            ["client_secret"] = _googleClientSecret,
            ["refresh_token"] = refreshToken,
            ["grant_type"] = "refresh_token",
        });

        var googleHttpClient = _httpClientFactory.CreateClient("GoogleOAuth");
        var res = await googleHttpClient.PostAsync("https://oauth2.googleapis.com/token", content);
        if (!res.IsSuccessStatusCode) return null;

        var body = await res.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(body);
        var root = doc.RootElement;

        var token = root.TryGetProperty("access_token", out var t) ? t.GetString() : null;
        var expiresIn = root.TryGetProperty("expires_in", out var e) ? e.GetInt32() : 3600;

        return string.IsNullOrEmpty(token) ? null : (token, expiresIn);
    }

    /// <summary>
    /// Update user metadata using Clerk's merge API (/users/{id}/metadata).
    /// This merges keys instead of overwriting the entire private_metadata object,
    /// so other keys in privateMetadata are preserved.
    /// </summary>
    private async Task UpdatePrivateMetadata(string userId, Dictionary<string, object> metadata)
    {
        var payload = JsonSerializer.Serialize(new { private_metadata = metadata });
        var content = new StringContent(payload, System.Text.Encoding.UTF8, "application/json");

        using var request = new HttpRequestMessage(HttpMethod.Patch, $"users/{userId}/metadata");
        request.Content = content;

        var httpClient = _httpClientFactory.CreateClient("Clerk");
        await httpClient.SendAsync(request);
    }

    private static string? GetString(Dictionary<string, JsonElement>? dict, string key)
    {
        if (dict == null || !dict.TryGetValue(key, out var el)) return null;
        return el.ValueKind == JsonValueKind.String ? el.GetString() : null;
    }

    private static long? GetLong(Dictionary<string, JsonElement>? dict, string key)
    {
        if (dict == null || !dict.TryGetValue(key, out var el)) return null;
        return el.ValueKind == JsonValueKind.Number ? el.GetInt64() : null;
    }

    public class ClerkUserResponse
    {
        [JsonPropertyName("public_metadata")]
        public Dictionary<string, JsonElement>? PublicMetadata { get; set; }

        [JsonPropertyName("private_metadata")]
        public Dictionary<string, JsonElement>? PrivateMetadata { get; set; }
    }

    public class ClerkUserMeta
    {
        public string? SubscriptionStatus { get; set; }
        public string? StripeCustomerId { get; set; }
    }
}
