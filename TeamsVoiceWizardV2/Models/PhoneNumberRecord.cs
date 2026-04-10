using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace TeamsVoiceWizardV2.Models;

public class PhoneNumberRecord : INotifyPropertyChanged
{
    // --- Read-only fields from Graph numberAssignment ---
    public string TelephoneNumber { get; init; } = "";
    public string NumberType { get; init; } = "";
    public string AssignmentStatus { get; init; } = "";
    public string? AssignmentTargetId { get; init; }

    // --- Resolved display values ---
    public string? AssignedUserDisplayName { get; set; }
    public string? AssignedUserUpn { get; set; }

    // --- Derived flags for UI binding ---
    public bool IsUnassigned => AssignmentStatus == "unassigned";
    public bool IsDirectRouting => NumberType == "directRouting";
    public bool IsUserAssigned => AssignmentStatus == "userAssigned";
    public bool CanAssignPolicies => IsDirectRouting && IsUserAssigned;

    // --- Side panel pending/editing state ---
    private string? _pendingUserUpn;
    private string? _pendingDialPlan;
    private string? _pendingVoiceRoutingPolicy;

    // Current policy assignments fetched from Graph — null means not yet fetched
    public string? CurrentDialPlan { get; set; }
    public string? CurrentVoiceRoutingPolicy { get; set; }
    public bool PoliciesFetched { get; set; }

    public string? PendingUserUpn
    {
        get => _pendingUserUpn;
        set { _pendingUserUpn = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsDirty)); }
    }

    public string? PendingDialPlan
    {
        get => _pendingDialPlan;
        set { _pendingDialPlan = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsDirty)); }
    }

    public string? PendingVoiceRoutingPolicy
    {
        get => _pendingVoiceRoutingPolicy;
        set { _pendingVoiceRoutingPolicy = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsDirty)); }
    }

    public bool IsDirty =>
        (_pendingUserUpn != null && _pendingUserUpn != AssignedUserUpn) ||
        (_pendingDialPlan != null) ||
        (_pendingVoiceRoutingPolicy != null);

    public void ClearPending()
    {
        _pendingUserUpn = null;
        _pendingDialPlan = null;
        _pendingVoiceRoutingPolicy = null;
        OnPropertyChanged(nameof(PendingUserUpn));
        OnPropertyChanged(nameof(PendingDialPlan));
        OnPropertyChanged(nameof(PendingVoiceRoutingPolicy));
        OnPropertyChanged(nameof(IsDirty));
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
