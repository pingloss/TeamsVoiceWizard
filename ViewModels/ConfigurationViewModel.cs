using System.Collections.ObjectModel;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.UI.Dispatching;
using TeamsVoiceWizard.Models;
using TeamsVoiceWizard.Services;

namespace TeamsVoiceWizard.ViewModels;

public partial class ConfigurationViewModel : ObservableObject
{
    private readonly ConfigurationHostServices _host;
    private readonly DispatcherQueue            _dispatcher;
    private readonly WizardState                _state = new();
    private          List<SkuRecord>            _skuChoices = [];

    // Buffered logging (prevents O(n²) string concatenation on rapid PS output)
    private readonly StringBuilder _logBuffer = new();
    private readonly StringBuilder _outBuffer = new();
    private          bool          _logFlushScheduled;
    private          bool          _outFlushScheduled;
    private const    int           MaxTextChars = 250_000;

    public ConfigurationViewModel(ConfigurationHostServices host)
    {
        _host       = host;
        _dispatcher = DispatcherQueue.GetForCurrentThread();
    }

    // ── Events for view to react to (scroll-to-bottom etc.) ──────────────────

    public event EventHandler? LogScrollRequested;
    public event EventHandler? OutputScrollRequested;

    // ── Form inputs ───────────────────────────────────────────────────────────

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsDomainsTxtEnabled))]
    private string _gammaEndpoint = "";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsDomainsTxtEnabled))]
    private string _siteName = "";

    [ObservableProperty]
    private string _country = "GB";

    [ObservableProperty]
    private bool _derivedTrunkModel = true;

    // ── Connection / PS-ready state ───────────────────────────────────────────

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsConnectTeamsEnabled))]
    [NotifyPropertyChangedFor(nameof(IsLoadSkusEnabled))]
    [NotifyPropertyChangedFor(nameof(IsCreateTestObjsEnabled))]
    private bool _graphConnected;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsApplyVoiceEnabled))]
    private bool _teamsConnected;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsConnectGraphEnabled))]
    [NotifyPropertyChangedFor(nameof(IsConnectTeamsEnabled))]
    [NotifyPropertyChangedFor(nameof(IsDomainsTxtEnabled))]
    [NotifyPropertyChangedFor(nameof(IsVerifyEnabled))]
    [NotifyPropertyChangedFor(nameof(IsLoadSkusEnabled))]
    [NotifyPropertyChangedFor(nameof(IsCreateTestObjsEnabled))]
    [NotifyPropertyChangedFor(nameof(IsCleanupEnabled))]
    [NotifyPropertyChangedFor(nameof(IsApplyVoiceEnabled))]
    private bool _isPsReady;

    // ── Busy state ────────────────────────────────────────────────────────────

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsConnectGraphEnabled))]
    [NotifyPropertyChangedFor(nameof(IsConnectTeamsEnabled))]
    private bool _isBusy;

    [ObservableProperty] private bool   _isBusyVisible;
    [ObservableProperty] private string _busyMessage = "";

    // ── SKU picker ────────────────────────────────────────────────────────────

    public ObservableCollection<string> SkuItems { get; } = [];

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsCreateTestObjsEnabled))]
    private int _selectedSkuIndex = -1;

    // ── Output / Log text (bound OneWay to read-only TextBoxes) ──────────────

    [ObservableProperty] private string _outputText = "";
    [ObservableProperty] private string _logText    = "";

    // ── Computed guardrail properties ─────────────────────────────────────────

    public bool IsConnectGraphEnabled  => IsPsReady && !IsBusy;
    public bool IsConnectTeamsEnabled  => IsPsReady && GraphConnected && !IsBusy;

    public bool IsDomainsTxtEnabled =>
        IsPsReady &&
        !string.IsNullOrWhiteSpace(GammaEndpoint) &&
        !string.IsNullOrWhiteSpace(SiteName);

    public bool IsVerifyEnabled      => IsPsReady && _state.Domains.Count > 0;
    public bool IsLoadSkusEnabled    => IsPsReady && GraphConnected;

    public bool IsCreateTestObjsEnabled =>
        IsPsReady && GraphConnected &&
        _state.Verification.Count > 0 && _state.Verification.All(v => v.Verified) &&
        SkuItems.Count > 0 && SelectedSkuIndex >= 0;

    public bool IsCleanupEnabled =>
        IsPsReady &&
        (_state.CreatedUsers.Count > 0 || _state.CreatedResourceAccounts.Count > 0);

    public bool IsApplyVoiceEnabled =>
        IsPsReady && TeamsConnected &&
        _state.Verification.Count > 0 && _state.Verification.All(v => v.Verified);

    // ── Public API called by MainWindow ───────────────────────────────────────

    /// <summary>Called by MainWindow once the PowerShell runspace has finished initialising.</summary>
    public void SetPowerShellReady(bool ready)
    {
        IsPsReady = ready;
    }

    // ── Logging (public so PhoneManagementHostServices can share the log box) ─

    public void AppendLog(string line)
    {
        lock (_logBuffer)
        {
            _logBuffer.AppendLine(line);
            if (_logBuffer.Length > MaxTextChars)
                _logBuffer.Remove(0, _logBuffer.Length - MaxTextChars);
            if (_logFlushScheduled) return;
            _logFlushScheduled = true;
        }
        _dispatcher.TryEnqueue(FlushLog);
    }

    private void FlushLog()
    {
        string text;
        lock (_logBuffer)
        {
            text               = _logBuffer.ToString();
            _logFlushScheduled = false;
        }
        LogText = text;
        LogScrollRequested?.Invoke(this, EventArgs.Empty);
    }

    private void AppendOutput(string text)
    {
        lock (_outBuffer)
        {
            _outBuffer.Append(text);
            if (_outBuffer.Length > MaxTextChars)
                _outBuffer.Remove(0, _outBuffer.Length - MaxTextChars);
            if (_outFlushScheduled) return;
            _outFlushScheduled = true;
        }
        _dispatcher.TryEnqueue(FlushOutput);
    }

    private void FlushOutput()
    {
        string text;
        lock (_outBuffer)
        {
            text               = _outBuffer.ToString();
            _outFlushScheduled = false;
        }
        OutputText = text;
        OutputScrollRequested?.Invoke(this, EventArgs.Empty);
    }

    // ── Busy helper ───────────────────────────────────────────────────────────

    private void SetBusy(bool busy, string message = "")
    {
        IsBusy        = busy;
        IsBusyVisible = busy;
        BusyMessage   = message;
        NotifyGuardrailsChanged();
    }

    private void NotifyGuardrailsChanged()
    {
        OnPropertyChanged(nameof(IsConnectGraphEnabled));
        OnPropertyChanged(nameof(IsConnectTeamsEnabled));
        OnPropertyChanged(nameof(IsDomainsTxtEnabled));
        OnPropertyChanged(nameof(IsVerifyEnabled));
        OnPropertyChanged(nameof(IsLoadSkusEnabled));
        OnPropertyChanged(nameof(IsCreateTestObjsEnabled));
        OnPropertyChanged(nameof(IsCleanupEnabled));
        OnPropertyChanged(nameof(IsApplyVoiceEnabled));
    }

    // ── Commands ──────────────────────────────────────────────────────────────

    [RelayCommand]
    private async Task ConnectGraphAsync()
    {
        if (!IsPsReady) return;

        SetBusy(true, "Connecting to Microsoft Graph...");
        try
        {
            AppendLog("Graph: Starting device-code login...");
            var scopes = new[]
            {
                "User.ReadWrite.All",
                "Domain.ReadWrite.All",
                "Organization.Read.All",
                "TeamsPolicyUserAssign.ReadWrite.All",
                "TeamsTelephoneNumber.ReadWrite.All",
                "TeamsUserConfiguration.Read.All"
            };
            await _host.ConnectGraphAsync(scopes);
            GraphConnected = await _host.IsGraphConnectedAsync();
            AppendLog(GraphConnected
                ? "Graph: Connected."
                : "Graph: Login completed but validation failed.");
        }
        catch (Exception ex)
        {
            AppendLog($"Graph: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }
    }

    [RelayCommand]
    private async Task ConnectTeamsAsync()
    {
        if (!IsPsReady) return;

        SetBusy(true, "Connecting to Microsoft Teams...");
        try
        {
            AppendLog("Teams: Starting authentication...");
            await _host.ConnectTeamsAsync();
            TeamsConnected = await _host.IsTeamsConnectedAsync();
            AppendLog(TeamsConnected
                ? "Teams: Connected."
                : "Teams: Connection failed.");
        }
        catch (Exception ex)
        {
            AppendLog($"Teams: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }
    }

    [RelayCommand]
    private async Task CreateDomainsAsync()
    {
        if (!IsPsReady) return;

        SetBusy(true, "Creating domains and generating TXT records...");
        try
        {
            var endpoint = GammaEndpoint.Trim();
            var site     = SiteName.Trim();
            var country  = Country.Trim();

            if (string.IsNullOrWhiteSpace(endpoint)) { AppendLog("PSTN Gateway(s) input is required."); return; }
            if (string.IsNullOrWhiteSpace(site))     { AppendLog("Site name is required."); return; }

            _state.GammaEndpoint = endpoint;
            _state.Site          = site;
            _state.Country       = country;

            AppendLog($"Creating voice domains from: {endpoint}");

            var (domains, derivedDetected, outputText) =
                await _host.CreateDomainsAndGetTxtAsync(endpoint, site, country);

            _state.Domains = domains;

            if (derivedDetected)
                DerivedTrunkModel = true;

            AppendOutput(outputText + Environment.NewLine);
            AppendLog($"Domains created: {string.Join(", ", domains)}");
            NotifyGuardrailsChanged();
        }
        catch (Exception ex)
        {
            AppendLog($"Domains/TXT: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }
    }

    [RelayCommand]
    private async Task VerifyDomainsAsync()
    {
        if (!IsPsReady) return;

        SetBusy(true, "Verifying domain ownership...");
        try
        {
            if (_state.Domains.Count == 0) { AppendLog("No domains to verify."); return; }

            AppendLog("Verifying domains...");
            _state.Verification = await _host.VerifyDomainsAsync();

            var table = string.Join(Environment.NewLine,
                _state.Verification.Select(v =>
                    $"  {v.Domain,-55} Verified={v.Verified,-5} {v.Error ?? ""}"));

            AppendOutput($"\nVerification:\n{table}\n");

            AppendLog(_state.Verification.All(v => v.Verified)
                ? "All domains verified."
                : $"Verification complete — {_state.Verification.Count(v => !v.Verified)} domain(s) failed.");

            NotifyGuardrailsChanged();
        }
        catch (Exception ex)
        {
            AppendLog($"Verify: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }
    }

    [RelayCommand]
    private async Task LoadLicenceSkusAsync()
    {
        if (!IsPsReady) return;

        SetBusy(true, "Loading licence SKUs...");
        try
        {
            if (!GraphConnected) { AppendLog("Graph not connected. Connect Graph first."); return; }

            AppendLog("Loading licence SKUs...");
            _skuChoices = await _host.LoadLicenseInventoryAsync(minFree: 2);

            SkuItems.Clear();
            SelectedSkuIndex = -1;

            if (_skuChoices.Count == 0)
            {
                AppendLog("No eligible SKUs found (none with ≥2 available seats).");
                return;
            }

            foreach (var sku in _skuChoices)
                SkuItems.Add(sku.ToString());

            SelectedSkuIndex = 0;
            AppendLog($"Loaded {_skuChoices.Count} eligible SKU(s).");
            NotifyGuardrailsChanged();
        }
        catch (Exception ex)
        {
            AppendLog($"Load SKUs: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }
    }

    [RelayCommand]
    private async Task CreateTestObjectsAsync()
    {
        if (!IsPsReady) return;

        SetBusy(true, "Creating test users and resource accounts...");
        try
        {
            if (!GraphConnected)                { AppendLog("Graph not connected."); return; }
            if (_state.Domains.Count == 0)      { AppendLog("No domains. Create domains first."); return; }
            if (SelectedSkuIndex < 0)           { AppendLog("No SKU selected."); return; }

            var sku = _skuChoices[SelectedSkuIndex];
            _state.ChosenSku = sku;

            if (sku.SkuId is null)              { AppendLog("Selected SKU has no SkuId in this tenant."); return; }
            if (sku.Available < 2)              { AppendLog("Selected SKU has fewer than 2 available seats."); return; }

            AppendLog($"Creating test objects with SKU: {sku.SkuPartNumber}");

            var (users, ras) = await _host.CreateTestObjectsAsync(sku, _state.Country);

            _state.CreatedUsers.AddRange(users.Except(_state.CreatedUsers));
            _state.CreatedResourceAccounts.AddRange(ras.Except(_state.CreatedResourceAccounts));

            AppendLog($"Created users: {users.Count}, resource accounts: {ras.Count}");
            NotifyGuardrailsChanged();
        }
        catch (Exception ex)
        {
            AppendLog($"Create test objects: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }
    }

    [RelayCommand]
    private async Task CleanupTestObjectsAsync()
    {
        if (!IsPsReady) return;

        SetBusy(true, "Cleaning up test objects...");
        try
        {
            if (_state.CreatedUsers.Count == 0 && _state.CreatedResourceAccounts.Count == 0)
            {
                AppendLog("No test objects recorded in this session.");
                return;
            }

            AppendLog("Cleaning up created test objects...");
            await _host.CleanupTestObjectsAsync(_state.CreatedUsers, _state.CreatedResourceAccounts);

            _state.CreatedUsers.Clear();
            _state.CreatedResourceAccounts.Clear();
            AppendLog("Cleanup complete.");
            NotifyGuardrailsChanged();
        }
        catch (Exception ex)
        {
            AppendLog($"Cleanup: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }
    }

    [RelayCommand]
    private async Task ApplyVoiceConfigurationAsync()
    {
        if (!IsPsReady) return;

        SetBusy(true, "Applying tenant-wide Teams Phone configuration...");
        try
        {
            if (!TeamsConnected)                    { AppendLog("Teams not connected. Connect Teams first."); return; }
            if (string.IsNullOrWhiteSpace(_state.Site)) { AppendLog("Site name is required."); return; }
            if (_state.Domains.Count == 0)          { AppendLog("No PSTN gateways. Create domains first."); return; }

            AppendLog("⚠ Applying tenant-wide Teams Phone configuration...");
            await _host.ApplyVoiceConfigurationAsync(_state.Site, _state.Country, DerivedTrunkModel);
        }
        catch (Exception ex)
        {
            AppendLog($"Apply voice config: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }
    }
}
