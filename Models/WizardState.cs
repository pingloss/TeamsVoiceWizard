namespace TeamsVoiceWizard.Models;

/// <summary>
/// C# mirror of the PowerShell New-TVWState pscustomobject.
/// Holds UI-visible state; the PowerShell runspace holds the
/// corresponding PS-side variables in parallel.
/// </summary>
public sealed class WizardState
{
    public string GammaEndpoint { get; set; } = string.Empty;
    public string Site { get; set; } = string.Empty;
    public string Country { get; set; } = "GB";
    public List<string> Domains { get; set; } = [];
    public List<VerificationResult> Verification { get; set; } = [];
    public List<SkuRecord> LicenseInventory { get; set; } = [];
    public SkuRecord? ChosenSku { get; set; }
    public List<string> CreatedUsers { get; set; } = [];
    public List<string> CreatedResourceAccounts { get; set; } = [];
    public bool DerivedTrunkModel { get; set; } = true;
}

public sealed record VerificationResult(string Domain, bool Verified, string? Error);

public sealed record SkuRecord(
    string Family,
    string Product,
    string SkuPartNumber,
    Guid? SkuId,
    int TotalEnabled,
    int Consumed,
    int Available,
    bool MeetsMinFree,
    bool ExistsInTenant)
{
    public override string ToString() =>
        $"{Product} | {SkuPartNumber} | Free:{Available} | {Family}";
}