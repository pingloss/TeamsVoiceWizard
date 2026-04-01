using Microsoft.UI;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using System.Diagnostics;
using System.Text;

using TeamsVoiceWizard.Models;
using TeamsVoiceWizard.Services;

using Windows.Graphics;
using Windows.Media.AppBroadcasting;
using WinRT.Interop;

namespace TeamsVoiceWizard;

public sealed partial class MainWindow : Window
{
    private PowerShellHost? _ps;
    private readonly WizardState _state = new();
    private readonly DispatcherQueue _dispatcher;

    private bool _uiReady;
    private bool _startupRan;
    private bool _psReady;

    private bool _graphConnected;
    private bool _teamsConnected;

    private List<SkuRecord> _skuChoices = [];

    // Buffered UI logging (prevents O(n²) Text concatenation)
    private readonly StringBuilder _logBuffer = new();
    private readonly StringBuilder _outBuffer = new();
    private bool _logFlushScheduled;
    private bool _outFlushScheduled;
    private const int MaxTextChars = 250_000; // cap buffers to avoid runaway memory

    public MainWindow()
    {
        InitializeComponent();

        Title = "Teams Voice Wizard";
        _dispatcher = DispatcherQueue.GetForCurrentThread();

        RootGrid.Loaded += RootGrid_Loaded;
        Closed += MainWindow_Closed;
    }

    private async void RootGrid_Loaded(object sender, RoutedEventArgs e)
    {
        if (_startupRan) return;
        _startupRan = true;

        _uiReady = true;

        TrySetInitialWindowSize(width: 1140, height: 880);

        // Sync initial state
        _state.Country = (CountryBox?.Text ?? "GB").Trim().ToUpperInvariant();
        _state.DerivedTrunkModel = (ChkDerivedTrunk?.IsChecked == true);

        AppendLog("Teams Voice Wizard started.");

        var sw = Stopwatch.StartNew();
        AppendLog($"[Perf] Startup: begin @ {sw.ElapsedMilliseconds}ms");

        await InitPowerShellAsync();

        AppendLog($"[Perf] Startup: PS init complete @ {sw.ElapsedMilliseconds}ms");

        UpdateGuardrails();
    }

    private void MainWindow_Closed(object sender, WindowEventArgs args)
    {
        try { _ps?.Dispose(); }
        catch { /* ignore */ }
        finally
        {
            _ps = null;
            _psReady = false;
        }
    }

    private void TrySetInitialWindowSize(int width, int height)
    {
        try
        {
            // WinUI Window sizing uses AppWindow APIs. [2](https://stackoverflow.com/questions/54563576/uwp-app-build-failure-microsoft-ui-xaml-markup-could-not-be-found)[3](https://www.thewindowsclub.com/process-exited-with-code-1)
            var hwnd = WindowNative.GetWindowHandle(this);
            var windowId = Win32Interop.GetWindowIdFromWindow(hwnd);
            var appWindow = AppWindow.GetFromWindowId(windowId);

            appWindow.Resize(new SizeInt32(width, height)); //[2](https://stackoverflow.com/questions/54563576/uwp-app-build-failure-microsoft-ui-xaml-markup-could-not-be-found)
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"TrySetInitialWindowSize failed: {ex}");
        }
    }

    private async Task InitPowerShellAsync()
    {
        try
        {
            var modulePath = Path.Combine(AppContext.BaseDirectory, "core", "TeamsVoiceWizard.Core.psm1");

            if (!File.Exists(modulePath))
            {
                AppendLog($"[Startup] PowerShell module not found: {modulePath}");
                AppendLog("[Startup] Ensure core\\TeamsVoiceWizard.Core.psm1 is copied to the output folder.");
                _psReady = false;
                return;
            }

            AppendLog("[Startup] Initialising PowerShell runspace...");

            // Create host off the UI thread so the window appears instantly.
            _ps = await Task.Run(() =>
                new PowerShellHost(modulePath, line =>
                {
                    _dispatcher.TryEnqueue(() => AppendLog(line));
                },
                diagnostics: false // toggle true when debugging module issues
                )
            );

            _psReady = true;

            // Validate connections (async, non-blocking)
            _graphConnected = await _ps.IsGraphConnectedAsync();
            _teamsConnected = await _ps.IsTeamsConnectedAsync();

            AppendLog("[Startup] PowerShell runspace ready.");
        }
        catch (Exception ex)
        {
            _psReady = false;
            _ps = null;
            AppendLog($"[Startup] PowerShell init failed: {ex.Message}");
        }
    }

    // -------------------------
    // Buffered logging
    // -------------------------

    private void AppendLog(string line)
    {
        if (!_uiReady)
        {
            Debug.WriteLine(line);
            return;
        }

        lock (_logBuffer)
        {
            _logBuffer.AppendLine(line);

            if (_logBuffer.Length > MaxTextChars)
                _logBuffer.Remove(0, _logBuffer.Length - MaxTextChars);

            if (_logFlushScheduled) return;
            _logFlushScheduled = true;
        }

        _dispatcher.TryEnqueue(FlushLogBuffer);
    }

    private void FlushLogBuffer()
    {
        if (LogBox is null) return;

        string text;
        lock (_logBuffer)
        {
            text = _logBuffer.ToString();
            _logFlushScheduled = false;
        }

        LogBox.Text = text;

        try { LogScroll?.ChangeView(null, double.MaxValue, null); }
        catch { /* ignore */ }
    }

    private void AppendOutput(string text)
    {
        if (!_uiReady)
        {
            Debug.WriteLine(text);
            return;
        }

        lock (_outBuffer)
        {
            _outBuffer.Append(text);

            if (_outBuffer.Length > MaxTextChars)
                _outBuffer.Remove(0, _outBuffer.Length - MaxTextChars);

            if (_outFlushScheduled) return;
            _outFlushScheduled = true;
        }

        _dispatcher.TryEnqueue(FlushOutBuffer);
    }

    private void FlushOutBuffer()
    {
        if (OutputBox is null) return;

        string text;
        lock (_outBuffer)
        {
            text = _outBuffer.ToString();
            _outFlushScheduled = false;
        }

        OutputBox.Text = text;

        try { OutputScroll?.ChangeView(null, double.MaxValue, null); }
        catch { /* ignore */ }
    }

    // -------------------------
    // Guardrails
    // -------------------------

    private bool EnsureUiWired()
    {
        return GammaBox is not null &&
               SiteBox is not null &&
               CountryBox is not null &&
               ChkDerivedTrunk is not null &&
               BtnConnectGraph is not null &&
               BtnConnectTeams is not null &&
               BtnDomainsTxt is not null &&
               BtnVerify is not null &&
               BtnLoadSkus is not null &&
               SkuDrop is not null &&
               BtnCreateTestObjs is not null &&
               BtnCleanup is not null &&
               BtnApplyVoice is not null;
    }

    private bool EnsurePowerShellReady(bool logIfNotReady = true)
    {
        if (_psReady && _ps is not null) return true;

        if (logIfNotReady)
            AppendLog("PowerShell host not ready (startup failed or module missing).");

        return false;
    }

    private void UpdateGuardrails()
    {
        if (!_uiReady) return;
        if (!EnsureUiWired()) return;

        BtnConnectGraph.IsEnabled = _psReady;
        BtnConnectTeams.IsEnabled = _psReady && _graphConnected;

        BtnDomainsTxt.IsEnabled =
            _psReady &&
            !string.IsNullOrWhiteSpace(GammaBox.Text) &&
            !string.IsNullOrWhiteSpace(SiteBox.Text);

        BtnVerify.IsEnabled = _psReady && _state.Domains.Count > 0;
        BtnLoadSkus.IsEnabled = _psReady && _graphConnected;

        var domainsVerified =
            _state.Verification.Count > 0 &&
            _state.Verification.All(v => v.Verified);

        var skuSelected = _skuChoices.Count > 0 && SkuDrop.SelectedIndex >= 0;

        BtnCreateTestObjs.IsEnabled = _psReady && _graphConnected && domainsVerified && skuSelected;

        BtnCleanup.IsEnabled =
            _psReady &&
            (_state.CreatedUsers.Count > 0 || _state.CreatedResourceAccounts.Count > 0);

        BtnApplyVoice.IsEnabled = _psReady && _teamsConnected && domainsVerified;
    }

    private void SetBusy(bool busy, string? message = null)
    {
        if (!_uiReady) return;

        StatusPanel.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
        BusyRing.IsActive = busy;
        StatusText.Text = message ?? string.Empty;
    }

    // -------------------------
    // Input handlers
    // -------------------------

    private void GammaBox_TextChanged(object sender, TextChangedEventArgs e) { if (_uiReady) UpdateGuardrails(); }
    private void SiteBox_TextChanged(object sender, TextChangedEventArgs e) { if (_uiReady) UpdateGuardrails(); }

    private void CountryBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        _state.Country = (CountryBox?.Text ?? "GB").Trim().ToUpperInvariant();
        if (_uiReady) UpdateGuardrails();
    }

    private void SkuDrop_SelectionChanged(object sender, SelectionChangedEventArgs e) { if (_uiReady) UpdateGuardrails(); }

    private void ChkDerivedTrunk_Checked(object sender, RoutedEventArgs e)
    {
        _state.DerivedTrunkModel = true;
        if (_uiReady) UpdateGuardrails();
    }

    private void ChkDerivedTrunk_Unchecked(object sender, RoutedEventArgs e)
    {
        _state.DerivedTrunkModel = false;
        if (_uiReady) UpdateGuardrails();
    }

    // -------------------------
    // Button handlers
    // -------------------------

    private async void BtnConnectGraph_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsurePowerShellReady()) return;

        BtnConnectGraph.IsEnabled = false;
        SetBusy(true, "Connecting to Microsoft Graph...");
        try
        {
            AppendLog("Graph: Starting device-code login...");
            var scopes = new[] { "User.ReadWrite.All", "Domain.ReadWrite.All", "Organization.Read.All" };
            await _ps!.ConnectGraphAsync(scopes);

            _graphConnected = await _ps.IsGraphConnectedAsync();
            AppendLog(_graphConnected ? "Graph: Connected." : "Graph: Login completed but validation failed.");
        }
        catch (Exception ex)
        {
            AppendLog($"Graph: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
            BtnConnectGraph.IsEnabled = true;
            UpdateGuardrails();
        }
    }

    private async void BtnConnectTeams_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsurePowerShellReady()) return;

        BtnConnectTeams.IsEnabled = false;
        SetBusy(true, "Connecting to Microsoft Teams...");
        try
        {
            AppendLog("Teams: Starting authentication...");
            await _ps!.ConnectTeamsAsync();

            _teamsConnected = await _ps.IsTeamsConnectedAsync();
            AppendLog(_teamsConnected ? "Teams: Connected." : "Teams: Connection failed.");
        }
        catch (Exception ex)
        {
            AppendLog($"Teams: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
            UpdateGuardrails();
        }
    }

    private async void BtnDomainsTxt_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsurePowerShellReady()) return;

        BtnDomainsTxt.IsEnabled = false;
        SetBusy(true, "Creating domains and generating TXT records...");
        try
        {
            var endpoint = (GammaBox?.Text ?? "").Trim();
            var site = (SiteBox?.Text ?? "").Trim();
            var country = (CountryBox?.Text ?? "GB").Trim();

            if (string.IsNullOrWhiteSpace(endpoint)) { AppendLog("PSTN Gateway(s) input is required."); return; }
            if (string.IsNullOrWhiteSpace(site)) { AppendLog("Site name is required."); return; }

            _state.GammaEndpoint = endpoint;
            _state.Site = site;
            _state.Country = country;

            AppendLog($"Creating voice domains from: {endpoint}");

            var (domains, derivedDetected, outputText) =
                await _ps!.CreateDomainsAndGetTxtAsync(endpoint, site, country);

            _state.Domains = domains;

            if (derivedDetected)
            {
                _state.DerivedTrunkModel = true;
                if (ChkDerivedTrunk is not null)
                    ChkDerivedTrunk.IsChecked = true;
            }

            AppendOutput(outputText + Environment.NewLine);
            AppendLog($"Domains created: {string.Join(", ", domains)}");
        }
        catch (Exception ex)
        {
            AppendLog($"Domains/TXT: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
            BtnDomainsTxt.IsEnabled = true;
            UpdateGuardrails();
        }
    }

    private async void BtnVerify_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsurePowerShellReady()) return;

        BtnVerify.IsEnabled = false;
        SetBusy(true, "Verifying domain ownership...");
        try
        {
            if (_state.Domains.Count == 0) { AppendLog("No domains to verify."); return; }

            AppendLog("Verifying domains...");
            _state.Verification = await _ps!.VerifyDomainsAsync();

            var table = string.Join(Environment.NewLine,
                _state.Verification.Select(v => $"  {v.Domain,-55} Verified={v.Verified,-5} {v.Error ?? ""}"));

            AppendOutput($"\nVerification:\n{table}\n");

            AppendLog(_state.Verification.All(v => v.Verified)
                ? "All domains verified."
                : $"Verification complete — {_state.Verification.Count(v => !v.Verified)} domain(s) failed.");
        }
        catch (Exception ex)
        {
            AppendLog($"Verify: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
            BtnVerify.IsEnabled = true;
            UpdateGuardrails();
        }
    }

    private async void BtnLoadSkus_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsurePowerShellReady()) return;

        BtnLoadSkus.IsEnabled = false;
        SetBusy(true, "Loading license SKUs...");
        try
        {
            if (!_graphConnected) { AppendLog("Graph not connected. Connect Graph first."); return; }

            AppendLog("Loading license SKUs...");
            _skuChoices = await _ps!.LoadLicenseInventoryAsync(minFree: 2);

            SkuDrop!.Items.Clear();

            if (_skuChoices.Count == 0)
            {
                AppendLog("No eligible SKUs found (none with ≥2 available seats).");
                SkuDrop.SelectedIndex = -1;
                return;
            }

            foreach (var sku in _skuChoices)
                SkuDrop.Items.Add(sku.ToString());

            SkuDrop.SelectedIndex = 0;
            AppendLog($"Loaded {_skuChoices.Count} eligible SKU(s).");
        }
        catch (Exception ex)
        {
            AppendLog($"Load SKUs: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
            BtnLoadSkus.IsEnabled = true;
            UpdateGuardrails();
        }
    }

    private async void BtnCreateTestObjs_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsurePowerShellReady()) return;

        BtnCreateTestObjs.IsEnabled = false;
        SetBusy(true, "Creating test users and resource accounts...");
        try
        {
            if (!_graphConnected) { AppendLog("Graph not connected."); return; }
            if (_state.Domains.Count == 0) { AppendLog("No domains. Create domains first."); return; }
            if (SkuDrop!.SelectedIndex < 0) { AppendLog("No SKU selected."); return; }

            var sku = _skuChoices[SkuDrop.SelectedIndex];
            _state.ChosenSku = sku;

            if (sku.SkuId is null) { AppendLog("Selected SKU has no SkuId in this tenant."); return; }
            if (sku.Available < 2) { AppendLog("Selected SKU has fewer than 2 available seats."); return; }

            AppendLog($"Creating test objects with SKU: {sku.SkuPartNumber}");

            var (users, ras) = await _ps!.CreateTestObjectsAsync(sku, _state.Country);

            _state.CreatedUsers.AddRange(users.Except(_state.CreatedUsers));
            _state.CreatedResourceAccounts.AddRange(ras.Except(_state.CreatedResourceAccounts));

            AppendLog($"Created users: {users.Count}, resource accounts: {ras.Count}");
        }
        catch (Exception ex)
        {
            AppendLog($"Create test objects: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
            BtnCreateTestObjs.IsEnabled = true;
            UpdateGuardrails();
        }
    }

    private async void BtnCleanup_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsurePowerShellReady()) return;

        BtnCleanup.IsEnabled = false;
        SetBusy(true, "Cleaning up test objects...");
        try
        {
            if (_state.CreatedUsers.Count == 0 && _state.CreatedResourceAccounts.Count == 0)
            {
                AppendLog("No test objects recorded in this session.");
                return;
            }

            AppendLog("Cleaning up created test objects...");
            await _ps!.CleanupTestObjectsAsync(_state.CreatedUsers, _state.CreatedResourceAccounts);

            _state.CreatedUsers.Clear();
            _state.CreatedResourceAccounts.Clear();
            AppendLog("Cleanup complete.");
        }
        catch (Exception ex)
        {
            AppendLog($"Cleanup: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
            UpdateGuardrails();
        }
    }

    private async void BtnApplyVoice_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsurePowerShellReady()) return;

        BtnApplyVoice.IsEnabled = false;
        SetBusy(true, "Applying tenant-wide Teams Phone configuration...");
        try
        {
            if (!_teamsConnected) { AppendLog("Teams not connected. Connect Teams first."); return; }
            if (string.IsNullOrWhiteSpace(_state.Site)) { AppendLog("Site name is required."); return; }
            if (_state.Domains.Count == 0) { AppendLog("No PSTN gateways. Create domains first."); return; }

            AppendLog("⚠ Applying tenant-wide Teams Phone configuration...");
            await _ps!.ApplyVoiceConfigurationAsync(_state.Site, _state.Country, _state.DerivedTrunkModel);
        }
        catch (Exception ex)
        {
            AppendLog($"Apply voice config: Error — {ex.Message}");
        }
        finally
        {
            SetBusy(false);
            UpdateGuardrails();
        }
    }
}