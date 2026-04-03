namespace TeamsVoiceWizard.Models;

public record PolicyEntry(string Id, string DisplayName)
{
    public override string ToString() => DisplayName;
}

public record UserEntry(string ObjectId, string Upn, string DisplayName)
{
    public override string ToString() => $"{DisplayName} ({Upn})";
}

public class PolicyCaches
{
    public IReadOnlyList<PolicyEntry> DialPlans { get; set; } = [];
    public IReadOnlyList<PolicyEntry> VoiceRoutingPolicies { get; set; } = [];
    public IReadOnlyList<UserEntry> LicensedUsers { get; set; } = [];

    public bool PoliciesLoaded => DialPlans.Count > 0 || VoiceRoutingPolicies.Count > 0;
    public bool UsersLoaded => LicensedUsers.Count > 0;
}