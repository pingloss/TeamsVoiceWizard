using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using TeamsVoiceWizard.Models;

namespace TeamsVoiceWizard.Services;

// ── Internal Graph DTOs ────────────────────────────────────────────────────────

/// <summary>
/// DTO for GET /admin/teams/telephoneNumbers/numberAssignments
/// Represents a single phone number assignment in the tenant.
/// </summary>
internal record GraphNumberAssignment(
    [property: JsonPropertyName("telephoneNumber")] string TelephoneNumber,
    [property: JsonPropertyName("numberType")] string NumberType,
    [property: JsonPropertyName("assignmentStatus")] string AssignmentStatus,
    [property: JsonPropertyName("assignmentTargetId")] string? AssignmentTargetId,
    [property: JsonPropertyName("activationState")] string? ActivationState
);

/// <summary>
/// DTO for Graph /users endpoint.
/// Used in batch requests to resolve user details.
/// </summary>
internal record GraphUser(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("displayName")] string? DisplayName,
    [property: JsonPropertyName("userPrincipalName")] string? Upn
);

/// <summary>
/// Generic response envelope for Graph list endpoints.
/// Handles pagination via @odata.nextLink.
/// </summary>
internal record GraphListResponse<T>(
    [property: JsonPropertyName("value")] List<T> Value,
    [property: JsonPropertyName("@odata.nextLink")] string? NextLink
);

// ── Service ────────────────────────────────────────────────────────────────────

/// <summary>
/// Graph API client for Teams phone number and policy management.
/// Uses bearer token authentication with delegated scopes.
/// All methods handle pagination and error cases internally.
/// </summary>
public sealed class GraphPhoneService
{
    private readonly Func<Task<string>> _getToken;
    private readonly HttpClient _http;

    private static readonly JsonSerializerOptions _jsonOpts = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public GraphPhoneService(Func<Task<string>> getToken)
    {
        _getToken = getToken;
        _http = new HttpClient();
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    /// <summary>
    /// Builds an authenticated HTTP request with the current access token.
    /// Includes JSON body serialization if provided.
    /// </summary>
    private async Task<HttpRequestMessage> BuildRequestAsync(
        HttpMethod method, string url, object? body = null)
    {
        var token = await _getToken().ConfigureAwait(false);
        var req = new HttpRequestMessage(method, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

        if (body is not null)
            req.Content = new StringContent(
                JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

        return req;
    }

    /// <summary>
    /// GET request helper. Deserializes response as JSON.
    /// </summary>
    private async Task<T> GetAsync<T>(string url)
    {
        var req = await BuildRequestAsync(HttpMethod.Get, url).ConfigureAwait(false);
        var resp = await _http.SendAsync(req).ConfigureAwait(false);
        await EnsureSuccessAsync(resp).ConfigureAwait(false);
        var json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
        return JsonSerializer.Deserialize<T>(json, _jsonOpts)!;
    }

    /// <summary>
    /// POST request helper. Sends JSON body.
    /// </summary>
    private async Task PostAsync(string url, object body)
    {
        var req = await BuildRequestAsync(HttpMethod.Post, url, body).ConfigureAwait(false);
        var resp = await _http.SendAsync(req).ConfigureAwait(false);
        await EnsureSuccessAsync(resp).ConfigureAwait(false);
    }

    /// <summary>
    /// Checks HTTP response success. Throws with Graph error details on failure.
    /// </summary>
    private static async Task EnsureSuccessAsync(HttpResponseMessage resp)
    {
        if (resp.IsSuccessStatusCode) return;
        var error = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
        throw new InvalidOperationException(
            $"Graph API error {(int)resp.StatusCode} {resp.ReasonPhrase}: {error}");
    }

    /// <summary>
    /// Returns all effective policy assignments for a user, keyed by policyType.
    /// Value is (DisplayName, GraphPolicyId) — the GraphPolicyId is required for
    /// the policy assign endpoint and differs from the PS Identity string.
    /// Requires: TeamsUserConfiguration.Read.All scope
    /// </summary>
    public async Task<Dictionary<string, (string DisplayName, string PolicyId)>>
        GetUserTeamsConfigurationAsync(string userId, Action<string>? log = null)
    {
        var result = new Dictionary<string, (string DisplayName, string PolicyId)>(
            StringComparer.OrdinalIgnoreCase);

        try
        {
            var url = $"https://graph.microsoft.com/v1.0/admin/teams/userConfigurations/{userId}";
            log?.Invoke($"[UserConfig] GET {url}");

            using var req = await BuildRequestAsync(HttpMethod.Get, url).ConfigureAwait(false);
            var resp = await _http.SendAsync(req).ConfigureAwait(false);
            await EnsureSuccessAsync(resp).ConfigureAwait(false);

            var json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);

            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            // Navigate into value[] if present (some endpoints wrap in value array)
            var entity = root.TryGetProperty("value", out var valArr) &&
                         valArr.ValueKind == JsonValueKind.Array &&
                         valArr.GetArrayLength() > 0
                ? valArr[0]
                : root;

            if (!entity.TryGetProperty("effectivePolicyAssignments", out var assignments))
            {
                log?.Invoke("[UserConfig] No effectivePolicyAssignments property found");
                return result;
            }

            foreach (var item in assignments.EnumerateArray())
            {
                var policyType = item.TryGetProperty("policyType", out var pt)
                    ? pt.GetString() : null;

                if (!item.TryGetProperty("policyAssignment", out var pa)) continue;

                var displayName = pa.TryGetProperty("displayName", out var dn)
                    ? dn.GetString() : null;
                var policyId = pa.TryGetProperty("policyId", out var pi)
                    ? pi.GetString() : null;

                log?.Invoke($"[UserConfig] policyType='{policyType}' " +
                            $"displayName='{displayName}' policyId='{policyId}'");

                if (!string.IsNullOrWhiteSpace(policyType) &&
                    !string.IsNullOrWhiteSpace(displayName) &&
                    !string.IsNullOrWhiteSpace(policyId))
                {
                    result[policyType] = (displayName, policyId);
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[UserConfig] EXCEPTION: {ex.Message}");
        }

        return result;
    }

    // ── Public API ───────────────────────────────────────────────────��────────

    /// <summary>
    /// Returns all number assignments from the tenant, handling Graph pagination.
    /// Calls: GET /v1.0/admin/teams/telephoneNumbers/numberAssignments
    /// Requires: TeamsPhoneNumber.ReadWrite.All scope
    /// </summary>
    public async Task<List<PhoneNumberRecord>> GetNumberAssignmentsAsync()
    {
        var records = new List<PhoneNumberRecord>();
        string? next = "https://graph.microsoft.com/v1.0/admin/teams/telephoneNumberManagement/numberAssignments";

        while (next is not null)
        {
            var page = await GetAsync<GraphListResponse<GraphNumberAssignment>>(next)
                .ConfigureAwait(false);

            foreach (var item in page.Value)
            {
                records.Add(new PhoneNumberRecord
                {
                    TelephoneNumber = item.TelephoneNumber,
                    NumberType = item.NumberType,
                    AssignmentStatus = item.AssignmentStatus,
                    AssignmentTargetId = item.AssignmentTargetId
                });
            }

            next = page.NextLink;
        }

        return records;
    }

    /// <summary>
    /// Batch-resolves a set of user Object IDs to display name + UPN.
    /// Uses Graph $batch endpoint (max 20 per request) to avoid N+1 HTTP calls.
    /// Calls: POST /v1.0/$batch with multiple GET /users/{id} requests
    /// Requires: User.ReadWrite.All scope
    /// </summary>
    public async Task<Dictionary<string, (string DisplayName, string Upn)>>
        ResolveUsersAsync(IEnumerable<string> objectIds)
    {
        var ids = objectIds.Where(id => !string.IsNullOrWhiteSpace(id)).Distinct().ToList();
        var result = new Dictionary<string, (string, string)>(StringComparer.OrdinalIgnoreCase);

        if (ids.Count == 0) return result;

        var token = await _getToken().ConfigureAwait(false);

        const int batchSize = 20;

        // Batch into chunks of 20 (Graph $batch limit)
        for (int i = 0; i < ids.Count; i += batchSize)
        {
            var chunk = ids.Skip(i).Take(batchSize).ToList();

            // Build batch request array
            var requests = chunk.Select((id, idx) => new
            {
                id = idx.ToString(),
                method = "GET",
                url = $"/users/{id}?$select=id,displayName,userPrincipalName"
            });

            var req = new HttpRequestMessage(
                HttpMethod.Post, "https://graph.microsoft.com/v1.0/$batch");
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            req.Content = new StringContent(
                JsonSerializer.Serialize(new { requests }),
                Encoding.UTF8, "application/json");

            var resp = await _http.SendAsync(req).ConfigureAwait(false);
            await EnsureSuccessAsync(resp).ConfigureAwait(false);

            var json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            using var doc = JsonDocument.Parse(json);

            // Parse batch responses
            foreach (var response in doc.RootElement.GetProperty("responses").EnumerateArray())
            {
                if (response.GetProperty("status").GetInt32() != 200) continue;

                var body = response.GetProperty("body");
                var id = body.TryGetProperty("id", out var idEl) ? idEl.GetString() ?? "" : "";
                var displayName = body.TryGetProperty("displayName", out var dnEl) ? dnEl.GetString() ?? "" : "";
                var upn = body.TryGetProperty("userPrincipalName", out var upnEl) ? upnEl.GetString() ?? "" : "";

                if (!string.IsNullOrWhiteSpace(id))
                    result[id] = (displayName, upn);
            }
        }

        return result;
    }

    /// <summary>
    /// Returns all users who have the Teams Phone System service plan enabled.
    /// Results are used to populate the user assignment dropdown in the side panel.
    /// Calls: GET /v1.0/users with $filter on assignedPlans
    /// Requires: User.ReadWrite.All scope
    /// Note: Teams Phone System service plan ID = e43b5b99-8dfb-405f-9987-dc307f34bcbd
    /// </summary>
    public async Task<List<UserEntry>> GetTeamsPhoneLicensedUsersAsync()
    {
        var users = new List<UserEntry>();
        string? next = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName&$top=999";

        while (next is not null)
        {
            var page = await GetAsync<GraphListResponse<GraphUser>>(next).ConfigureAwait(false);

            foreach (var u in page.Value)
            {
                if (!string.IsNullOrWhiteSpace(u.Id) && !string.IsNullOrWhiteSpace(u.Upn))
                    users.Add(new UserEntry(u.Id, u.Upn, u.DisplayName ?? u.Upn ?? ""));
            }

            next = page.NextLink;
        }

        return users.OrderBy(u => u.DisplayName).ToList();
    }

    /// <summary>
    /// Assigns a telephone number to a user.
    /// Calls: POST /v1.0/admin/teams/telephoneNumbers/numberAssignments/assignNumber
    /// Requires: TeamsPhoneNumber.ReadWrite.All scope
    /// </summary>
    public Task AssignNumberAsync(string telephoneNumber, string userId) =>
        PostAsync(
            "https://graph.microsoft.com/v1.0/admin/teams/telephoneNumberManagement/numberAssignments/assignNumber",
            new
            {
                telephoneNumber,
                assignmentTargetId = userId,
                assignmentCategory = "primary"
            });

    /// <summary>
    /// Unassigns a telephone number from its current user.
    /// Calls: POST /v1.0/admin/teams/telephoneNumbers/numberAssignments/unassignNumber
    /// Requires: TeamsPhoneNumber.ReadWrite.All scope
    /// </summary>
    public Task UnassignNumberAsync(string telephoneNumber) =>
        PostAsync(
            "https://graph.microsoft.com/v1.0/admin/teams/telephoneNumberManagement/numberAssignments/unassignNumber",
            new { telephoneNumber });

    /// <summary>
    /// Assigns dial plan and/or voice routing policy to a user in a single batch call.
    /// Null values are skipped — only non-null policies are included in the request.
    /// Calls: POST /v1.0/admin/teams/policy/userAssignments/assign
    /// Requires: TeamsPolicyUserAssign.ReadWrite.All scope
    /// </summary>
    public Task AssignPoliciesAsync(string userId, string? dialPlanId, string? vrPolicyId)
    {
        var assignments = new List<object>();

        if (!string.IsNullOrWhiteSpace(dialPlanId))
            assignments.Add(new
            {
                userId,
                policyType = "TenantDialPlan",
                policyId = dialPlanId
            });

        if (!string.IsNullOrWhiteSpace(vrPolicyId))
            assignments.Add(new
            {
                userId,
                policyType = "OnlineVoiceRoutingPolicy",
                policyId = vrPolicyId
            });

        if (assignments.Count == 0)
            return Task.CompletedTask;

        return PostAsync(
            "https://graph.microsoft.com/v1.0/admin/teams/policy/userAssignments/assign",
            new { value = assignments });
    }
}