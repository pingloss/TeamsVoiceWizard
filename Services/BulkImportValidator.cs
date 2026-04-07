using System.Text.RegularExpressions;
using TeamsVoiceWizard.Models;

namespace TeamsVoiceWizard.Services;

/// <summary>
/// Two-phase validator for bulk import rows.
///
/// Phase 1 – Structural (no network):
///   • PhoneNumber present + E.164 format
///   • UserPrincipalName present
///   • Direct Routing rows require DialPlan and VoiceRoutingPolicy
///
/// Phase 2 – Tenant (Graph):
///   • PhoneNumber exists in tenant inventory
///   • UPN resolves to a real user
///   • User has Teams Phone licence
///   • Dial plan / VRP names exist in cached policy lists
///   • Resolves Graph policy IDs needed for the apply step
/// </summary>
public static class BulkImportValidator
{
    private static readonly Regex E164Pattern =
        new(@"^\+[1-9]\d{6,14}$", RegexOptions.Compiled);

    // ── Phase 1: Structural ───────────────────────────────────────────────────

    /// <summary>
    /// Validates structure only (no network). Sets Status = Error for any structural
    /// problems, Status = Pending for rows that need Phase 2 validation.
    /// Returns true if all rows pass structural validation.
    /// </summary>
    public static bool ValidateStructure(IEnumerable<BulkImportRow> rows)
    {
        bool allOk = true;

        foreach (var row in rows)
        {
            var errors = new List<string>();

            if (string.IsNullOrWhiteSpace(row.PhoneNumber))
                errors.Add("PhoneNumber is required.");
            else if (!E164Pattern.IsMatch(row.PhoneNumber))
                errors.Add($"'{row.PhoneNumber}' is not a valid E.164 number (must start with + followed by 7–15 digits).");

            if (string.IsNullOrWhiteSpace(row.UserPrincipalName))
                errors.Add("UserPrincipalName is required.");
            else if (!row.UserPrincipalName.Contains('@'))
                errors.Add($"'{row.UserPrincipalName}' does not look like a valid UPN.");

            if (errors.Count > 0)
            {
                row.Status            = BulkImportStatus.Error;
                row.ValidationMessage = string.Join(" ", errors);
                allOk = false;
            }
            else
            {
                row.Status            = BulkImportStatus.Pending;
                row.ValidationMessage = "";
            }
        }

        return allOk;
    }

    // ── Phase 2: Tenant ───────────────────────────────────────────────────────

    /// <summary>
    /// Validates all structurally-valid rows against live tenant data.
    /// Must be called from a background thread (makes Graph calls).
    /// Reports progress via <paramref name="onProgress"/>.
    /// </summary>
    public static async Task ValidateTenantAsync(
        IReadOnlyList<BulkImportRow>           rows,
        GraphPhoneService                       graph,
        IReadOnlyList<PhoneNumberRecord>        tenantNumbers,
        IReadOnlyList<PolicyEntry>              dialPlans,
        IReadOnlyList<PolicyEntry>              vrPolicies,
        Dictionary<string, string>              graphPolicyIdCache,
        Action<BulkImportRow>?                  onProgress,
        CancellationToken                       ct = default)
    {
        // Only validate rows that passed structural checks
        var candidates = rows.Where(r => r.Status == BulkImportStatus.Pending).ToList();
        if (candidates.Count == 0) return;

        // ── Step A: Build tenant number lookup ───────────────────────────────
        var numberLookup = tenantNumbers.ToDictionary(
            n => n.TelephoneNumber.TrimStart('+'),
            n => n,
            StringComparer.OrdinalIgnoreCase);

        // Also index with + prefix
        foreach (var n in tenantNumbers)
            numberLookup.TryAdd(n.TelephoneNumber.TrimStart('+'), n);

        // ── Step B: Batch-resolve UPNs → user details + licensing ───────────
        var uniqueUpns = candidates
            .Select(r => r.UserPrincipalName)
            .Where(u => !string.IsNullOrWhiteSpace(u))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        ct.ThrowIfCancellationRequested();

        var userLookup = await graph
            .ResolveUsersByUpnBatchAsync(uniqueUpns, ct)
            .ConfigureAwait(false);

        // ── Step C: Build policy name lookups ────────────────────────────────
        var dpByName = dialPlans.ToDictionary(
            p => StripTag(p.DisplayName),
            p => p,
            StringComparer.OrdinalIgnoreCase);

        var vrpByName = vrPolicies.ToDictionary(
            p => StripTag(p.DisplayName),
            p => p,
            StringComparer.OrdinalIgnoreCase);

        // ── Step D: Validate each row ─────────────────────────────────────────
        foreach (var row in candidates)
        {
            ct.ThrowIfCancellationRequested();
            row.Status = BulkImportStatus.Validating;
            onProgress?.Invoke(row);

            var errors = new List<string>();

            // Number existence
            var normalised = row.PhoneNumber.TrimStart('+');
            if (!numberLookup.TryGetValue(normalised, out var tenantNumber))
            {
                errors.Add($"Number {row.PhoneNumber} not found in tenant.");
            }
            else
            {
                row.NumberType = tenantNumber.NumberType;
            }

            // User existence
            if (!userLookup.TryGetValue(row.UserPrincipalName, out var userData))
            {
                errors.Add($"User '{row.UserPrincipalName}' not found in tenant.");
            }
            else
            {
                row.ResolvedUserId      = userData.ObjectId;
                row.ResolvedDisplayName = userData.DisplayName;

                if (!userData.HasTeamsPhoneLicence)
                    errors.Add($"User '{row.UserPrincipalName}' does not have a Teams Phone licence.");
            }

            // Policy requirements for Direct Routing
            if (row.NumberType is not null &&
                string.Equals(row.NumberType, "directRouting", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrWhiteSpace(row.DialPlan))
                    errors.Add("DialPlan is required for Direct Routing numbers.");

                if (string.IsNullOrWhiteSpace(row.VoiceRoutingPolicy))
                    errors.Add("VoiceRoutingPolicy is required for Direct Routing numbers.");
            }

            // Dial plan existence + ID resolution (if provided)
            if (!string.IsNullOrWhiteSpace(row.DialPlan))
            {
                var dpName = StripTag(row.DialPlan);

                if (!dpByName.TryGetValue(dpName, out var dpEntry))
                {
                    errors.Add($"Dial plan '{row.DialPlan}' not found in tenant.");
                }
                else
                {
                    // Try to resolve Graph policy ID
                    var cacheKey = $"TenantDialPlan:{dpEntry.DisplayName}";
                    row.ResolvedDialPlanGraphId = graphPolicyIdCache.TryGetValue(cacheKey, out var dpId)
                        ? dpId : dpEntry.Id; // fall back to PS Identity string
                }
            }

            // VRP existence + ID resolution (if provided)
            if (!string.IsNullOrWhiteSpace(row.VoiceRoutingPolicy))
            {
                var vrpName = StripTag(row.VoiceRoutingPolicy);

                if (!vrpByName.TryGetValue(vrpName, out var vrpEntry))
                {
                    errors.Add($"Voice routing policy '{row.VoiceRoutingPolicy}' not found in tenant.");
                }
                else
                {
                    var cacheKey = $"OnlineVoiceRoutingPolicy:{vrpEntry.DisplayName}";
                    row.ResolvedVrpGraphId = graphPolicyIdCache.TryGetValue(cacheKey, out var vrpId)
                        ? vrpId : vrpEntry.Id;
                }
            }

            // Finalise
            if (errors.Count > 0)
            {
                row.Status            = BulkImportStatus.Error;
                row.ValidationMessage = string.Join(" ", errors);
            }
            else
            {
                row.Status            = BulkImportStatus.Valid;
                row.ValidationMessage = "";
            }

            onProgress?.Invoke(row);
        }
    }

    private static string StripTag(string value) =>
        value.StartsWith("Tag:", StringComparison.OrdinalIgnoreCase) ? value[4..] : value;
}
