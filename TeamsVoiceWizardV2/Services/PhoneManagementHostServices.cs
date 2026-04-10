using TeamsVoiceWizardV2.Models;

namespace TeamsVoiceWizardV2.Services;

/// <summary>
/// Bridges MainWindow (PowerShell host, Graph token, logging) into phone management MVVM.
/// </summary>
public sealed record PhoneManagementHostServices(
    Func<bool> IsGraphConnected,
    Func<bool, bool> TryEnsurePowerShell,
    Func<Task<string>> GetGraphAccessTokenAsync,
    Func<Task<List<PolicyEntry>>> GetDialPlansAsync,
    Func<Task<List<PolicyEntry>>> GetVoiceRoutingPoliciesAsync,
    Action<string> Log);
