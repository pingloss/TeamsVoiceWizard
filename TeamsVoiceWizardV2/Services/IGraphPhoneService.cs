using TeamsVoiceWizardV2.Models;

namespace TeamsVoiceWizardV2.Services;

/// <summary>
/// Abstraction over Graph API calls for Teams phone number and policy management.
/// Allows mock injection in tests and the prototype.
/// </summary>
public interface IGraphPhoneService
{
    Task<List<PhoneNumberRecord>> GetNumberAssignmentsAsync();

    Task<Dictionary<string, (string DisplayName, string Upn)>>
        ResolveUsersAsync(IEnumerable<string> objectIds);

    Task<List<UserEntry>> GetTeamsPhoneLicensedUsersAsync();

    Task<Dictionary<string, (string DisplayName, string PolicyId)>>
        GetUserTeamsConfigurationAsync(string userId, Action<string>? log = null);

    Task AssignNumberAsync(string telephoneNumber, string userId);

    Task UnassignNumberAsync(string telephoneNumber);

    Task AssignPoliciesAsync(string userId, string? dialPlanId, string? vrPolicyId);
}
