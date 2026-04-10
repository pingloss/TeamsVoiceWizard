using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace TeamsVoiceWizardV2.Models;

public enum BulkImportStatus { Pending, Validating, Valid, Error, Applied }

/// <summary>
/// Represents a single row parsed from a bulk import CSV or XLSX file.
/// Carries both the raw input values and all validation/resolution results.
/// </summary>
public class BulkImportRow : INotifyPropertyChanged
{
    // ── Raw input from file ───────────────────────────────────────────────────

    public int    RowNumber          { get; init; }
    public string PhoneNumber        { get; set; } = "";
    public string UserPrincipalName  { get; set; } = "";
    public string? DialPlan          { get; set; }
    public string? VoiceRoutingPolicy { get; set; }

    // ── Resolved at validation time ───────────────────────────────────────────

    public string? ResolvedUserId          { get; set; }
    public string? ResolvedDisplayName     { get; set; }
    public string? NumberType              { get; set; }  // "directRouting", "operatorConnect", "callingPlan"

    /// <summary>Graph policyId for the dial plan — resolved during validation.</summary>
    public string? ResolvedDialPlanGraphId      { get; set; }

    /// <summary>Graph policyId for the voice routing policy — resolved during validation.</summary>
    public string? ResolvedVrpGraphId           { get; set; }

    // ── Validation status ─────────────────────────────────────────────────────

    private BulkImportStatus _status = BulkImportStatus.Pending;
    public BulkImportStatus Status
    {
        get => _status;
        set
        {
            _status = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(StatusText));
            OnPropertyChanged(nameof(IsError));
            OnPropertyChanged(nameof(IsValid));
        }
    }

    private string _validationMessage = "";
    public string ValidationMessage
    {
        get => _validationMessage;
        set { _validationMessage = value; OnPropertyChanged(); }
    }

    public string StatusText => Status switch
    {
        BulkImportStatus.Pending    => "Pending",
        BulkImportStatus.Validating => "Validating…",
        BulkImportStatus.Valid      => "✓ Valid",
        BulkImportStatus.Error      => "✗ Error",
        BulkImportStatus.Applied    => "✓ Applied",
        _                           => ""
    };

    public bool IsError => Status == BulkImportStatus.Error;
    public bool IsValid => Status == BulkImportStatus.Valid;

    // ── Helpers ───────────────────────────────────────────────────────────────

    public bool RequiresPolicies =>
        string.Equals(NumberType, "directRouting", StringComparison.OrdinalIgnoreCase);

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
