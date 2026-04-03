using Microsoft.UI;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Text;
using Microsoft.UI.Xaml.Data;

using TeamsVoiceWizard.Models;
using TeamsVoiceWizard.Services;

using Windows.Graphics;
using WinRT.Interop;

using CommunityToolkit.WinUI.UI.Controls;




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

    // Phone management
    private GraphPhoneService? _graphPhone;
    private readonly ObservableCollection<PhoneNumberRecord> _phoneRecords = new();
    private readonly PolicyCaches _policyCaches = new();

    private bool _isLoadingDialPlans = false;
    private bool _isLoadingVRPolicies = false;
    private Dictionary<string, string> _usersCache = new();



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

        NumbersGrid.ItemsSource = _phoneRecords;
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

        if (PhoneTab is not null)
            PhoneTab.IsEnabled = _graphConnected;
    }

    private void SetBusy(bool busy, string? message = null)
    {
        if (!_uiReady) return;

        StatusPanel.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
        BusyRing.IsActive = busy;
        StatusText.Text = message ?? string.Empty;
    }

    private void SetPhoneBusy(bool busy, string? message = null)
    {
        if (!_uiReady) return;
        PhoneBusyRing.IsActive = busy;
        PhoneStatusText.Text = message ?? string.Empty;
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
            var scopes = new[] { "User.ReadWrite.All", "Domain.ReadWrite.All", "Organization.Read.All", "TeamsPolicyUserAssign.ReadWrite.All", "TeamsTelephoneNumber.ReadWrite.All" };
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

    private async void BtnLoadNumbers_Click(object sender, RoutedEventArgs e)
    {
        if (!_graphConnected) { AppendLog("Phone Management: Graph not connected."); return; }

        _graphPhone ??= new GraphPhoneService(() => _ps!.GetGraphAccessTokenAsync());

        BtnLoadNumbers.IsEnabled = false;
        SetPhoneBusy(true, "Loading phone numbers...");
        try
        {
            _phoneRecords.Clear();

            var records = await _graphPhone.GetNumberAssignmentsAsync();

            var assignedIds = records
                .Where(r => !string.IsNullOrWhiteSpace(r.AssignmentTargetId))
                .Select(r => r.AssignmentTargetId!)
                .Distinct().ToList();

            SetPhoneBusy(true, $"Resolving {assignedIds.Count} user(s)...");
            var resolved = assignedIds.Count > 0
                ? await _graphPhone.ResolveUsersAsync(assignedIds)
                : new Dictionary<string, (string DisplayName, string Upn)>();

            foreach (var r in records)
            {
                if (r.AssignmentTargetId is not null &&
                    resolved.TryGetValue(r.AssignmentTargetId, out var user))
                {
                    r.AssignedUserDisplayName = user.DisplayName;
                    r.AssignedUserUpn = user.Upn;
                }
                _phoneRecords.Add(r);
            }

            NumbersGrid.ItemsSource = _phoneRecords;
            AppendLog($"Phone Management: Loaded {_phoneRecords.Count} number(s).");
        }
        catch (Exception ex)
        {
            AppendLog($"Phone Management: Load failed — {ex.Message}");
        }
        finally
        {
            BtnLoadNumbers.IsEnabled = true;
            SetPhoneBusy(false);
        }
    }

    

    // Add these event handlers and methods to MainWindow.xaml.cs
    // Insert after the existing BtnLoadNumbers_Click and NumbersGrid_SelectionChanged methods

    // ════════════════════════════════════════════════════════════════════════════════
    // SIDE PANEL MANAGEMENT - Populate, populate, select handlers
    // ════════════════════════════════════════════════════════════════════════════════

    private PhoneNumberRecord? _currentlySelectedRecord;
    private Dictionary<string, string> _usersCache = new();  // userId -> displayName cache
    private bool _isLoadingUsers;

    private void NumbersGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        var selected = NumbersGrid.SelectedItem as PhoneNumberRecord;

        if (selected is null)
        {
            SidePanel.Visibility = Visibility.Collapsed;
            _currentlySelectedRecord = null;
            return;
        }

        _currentlySelectedRecord = selected;
        SidePanel.Visibility = Visibility.Visible;
        PopulateSidePanel(selected);
    }

    /// <summary>
    /// Populates the side panel with the selected record's data.
    /// Section A (Number Info) always shows.
    /// Section B (User Assignment) always shows.
    /// Section C (Policies) only shows when Direct Routing + assigned to user.
    /// </summary>
    private void PopulateSidePanel(PhoneNumberRecord record)
    {
        if (!_uiReady) return;

        // ══ SECTION A: Number Info (always visible) ══
        TxtPhoneNumber.Text = record.TelephoneNumber ?? "(unknown)";
        TxtNumberType.Text = record.NumberType ?? "(unknown)";
        TxtStatus.Text = record.AssignmentStatus ?? "(unknown)";
        TxtCurrentUser.Text = record.AssignedUserDisplayName ?? "(unassigned)";

        // ══ SECTION B: User Assignment ══
        PopulateUserComboBox(record);

        // ══ SECTION C: Policy Assignment Visibility ══
        // Only show policy section when: Direct Routing AND assigned to a user
        PolicyAssignmentSection.Visibility = record.CanAssignPolicies
            ? Visibility.Visible
            : Visibility.Collapsed;

        // Pre-populate policy ComboBoxes if assigned
        if (record.CanAssignPolicies)
        {
            PopulatePolicyComboBoxes(record);
        }

        // Reset the Apply button (no changes yet)
        BtnApplySingleNumber.IsEnabled = false;
    }

    /// <summary>
    /// Populate the User ComboBox. Called on demand when user first clicks the dropdown.
    /// </summary>
    private async void PopulateUserComboBox(PhoneNumberRecord record)
    {
        // If already populated from cache, just use it
        if (UserComboBox.Items?.Count > 0)
        {
            return;
        }

        // On first click, load the users list
        // We'll wire this up in the ComboBox's event handler
    }

    /// <summary>
    /// Called when the User ComboBox is opened (first time).
    /// Fetches licensed users from Graph on demand.
    /// </summary>
    private async void UserComboBox_DropDownOpened(object sender, object e)
    {
        if (_isLoadingUsers || UserComboBox.Items?.Count > 0)
        {
            return;
        }

        _isLoadingUsers = true;
        UserLoadingRing.IsActive = true;

        try
        {
            // If users cache is empty, fetch from Graph
            if (_usersCache.Count == 0)
            {
                var licensedUsers = await _graphPhone!.GetTeamsPhoneLicensedUsersAsync();

                foreach (var (userId, displayName, upn) in licensedUsers)
                {
                    _usersCache[userId] = displayName;
                    UserComboBox.Items?.Add(new ComboBoxItem { Content = displayName, Tag = upn });
                }

                AppendLog($"Phone Management: Loaded {_usersCache.Count} licensed Teams Phone user(s).");
            }
            else
            {
                // Already cached, just add to ComboBox if not there
                foreach (var kvp in _usersCache)
                {
                    if (!UserComboBox.Items!.Cast<ComboBoxItem>().Any(i => i.Content?.ToString() == kvp.Value))
                    {
                        UserComboBox.Items!.Add(new ComboBoxItem { Content = kvp.Value, Tag = kvp.Key });
                    }
                }
            }
        }
        catch (Exception ex)
        {
            AppendLog($"Phone Management: Failed to load users — {ex.Message}");
        }
        finally
        {
            UserLoadingRing.IsActive = false;
            _isLoadingUsers = false;
        }
    }

    private async void DialPlanComboBox_DropDownOpened(object sender, object e)
    {
        if (_isLoadingDialPlans || DialPlanComboBox.Items?.Count > 0)
            return;

        _isLoadingDialPlans = true;
        DialPlanLoadingRing.IsActive = true;

        try
        {
            if (!EnsurePowerShellReady()) return;

            var dialPlans = await _ps!.RunScalarAsync<List<string>>("Get-CsTenantDialPlan | Select-Object -ExpandProperty Identity");

            if (dialPlans == null || dialPlans.Count == 0)
            {
                AppendLog("Phone Management: No dial plans found.");
                return;
            }

            foreach (var plan in dialPlans)
            {
                DialPlanComboBox.Items?.Add(new ComboBoxItem { Content = plan });
            }

            AppendLog($"Phone Management: Loaded {dialPlans.Count} dial plan(s).");
        }
        catch (Exception ex)
        {
            AppendLog($"Phone Management: Failed to load dial plans — {ex.Message}");
        }
        finally
        {
            DialPlanLoadingRing.IsActive = false;
            _isLoadingDialPlans = false;
        }
    }

    private async void VoiceRoutingPolicyComboBox_DropDownOpened(object sender, object e)
    {
        if (_isLoadingVRPolicies || VoiceRoutingPolicyComboBox.Items?.Count > 0)
            return;

        _isLoadingVRPolicies = true;
        VRPolicyLoadingRing.IsActive = true;

        try
        {
            if (!EnsurePowerShellReady()) return;

            var policies = await _ps!.RunScalarAsync<List<string>>("Get-CsOnlineVoiceRoutingPolicy | Select-Object -ExpandProperty Identity");

            if (policies == null || policies.Count == 0)
            {
                AppendLog("Phone Management: No voice routing policies found.");
                return;
            }

            foreach (var policy in policies)
            {
                VoiceRoutingPolicyComboBox.Items?.Add(new ComboBoxItem { Content = policy });
            }

            AppendLog($"Phone Management: Loaded {policies.Count} voice routing policy(s).");
        }
        catch (Exception ex)
        {
            AppendLog($"Phone Management: Failed to load voice routing policies — {ex.Message}");
        }
        finally
        {
            VRPolicyLoadingRing.IsActive = false;
            _isLoadingVRPolicies = false;
        }
    }

    private void UserComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_currentlySelectedRecord is null) return;

        var selectedItem = UserComboBox.SelectedItem as ComboBoxItem;
        if (selectedItem is null)
        {
            return;
        }

        var selectedUserUpn = selectedItem.Tag?.ToString();
        _currentlySelectedRecord.PendingUserUpn = selectedUserUpn;

        // Mark as dirty and enable the Apply button
        UpdateApplyButtonState();
    }

    /// <summary>
    /// Populate policy ComboBoxes from the cached PS-fetched policies.
    /// </summary>
    private void PopulatePolicyComboBoxes(PhoneNumberRecord record)
    {
        // Clear existing items
        DialPlanComboBox.Items?.Clear();
        VoiceRoutingPolicyComboBox.Items?.Clear();

        // Add "(None / Keep Current)" option
        DialPlanComboBox.Items?.Add(new ComboBoxItem { Content = "(None - Keep Current)", Tag = null });
        VoiceRoutingPolicyComboBox.Items?.Add(new ComboBoxItem { Content = "(None - Keep Current)", Tag = null });

        // Add policies from cache
        foreach (var policy in _policyCaches.DialPlans)
        {
            DialPlanComboBox.Items?.Add(new ComboBoxItem { Content = policy.DisplayName, Tag = policy.Id });
        }

        foreach (var policy in _policyCaches.VoiceRoutingPolicies)
        {
            VoiceRoutingPolicyComboBox.Items?.Add(new ComboBoxItem { Content = policy.DisplayName, Tag = policy.Id });
        }
    }

    private void DialPlanComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_currentlySelectedRecord is null) return;

        var selectedItem = DialPlanComboBox.SelectedItem as ComboBoxItem;
        _currentlySelectedRecord.PendingDialPlan = selectedItem?.Tag?.ToString();

        UpdateApplyButtonState();
    }

    private void VoiceRoutingPolicyComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_currentlySelectedRecord is null) return;

        var selectedItem = VoiceRoutingPolicyComboBox.SelectedItem as ComboBoxItem;
        _currentlySelectedRecord.PendingVoiceRoutingPolicy = selectedItem?.Tag?.ToString();

        UpdateApplyButtonState();
    }

    /// <summary>
    /// Enable Apply button only when there are unsaved changes (IsDirty).
    /// </summary>
    private void UpdateApplyButtonState()
    {
        if (_currentlySelectedRecord is null)
        {
            BtnApplySingleNumber.IsEnabled = false;
            return;
        }

        BtnApplySingleNumber.IsEnabled = _currentlySelectedRecord.IsDirty;
    }

    /// <summary>
    /// Apply changes to the currently selected record via Graph API.
    /// </summary>
    private async void BtnApplySingleNumber_Click(object sender, RoutedEventArgs e)
    {
        if (_currentlySelectedRecord is null) return;

        BtnApplySingleNumber.IsEnabled = false;
        SetPhoneBusy(true, "Applying changes...");
        SidePanelStatusText.Text = "";
        SidePanelStatusText.Visibility = Visibility.Collapsed;

        try
        {
            // 1. Assign/unassign number if user selection changed
            if (_currentlySelectedRecord.PendingUserUpn != _currentlySelectedRecord.AssignedUserUpn)
            {
                SetPhoneBusy(true, $"Assigning number to user...");

                if (string.IsNullOrWhiteSpace(_currentlySelectedRecord.PendingUserUpn))
                {
                    // Unassign
                    await _graphPhone!.UnassignNumberAsync(_currentlySelectedRecord.TelephoneNumber);
                    AppendLog($"Phone Management: Unassigned {_currentlySelectedRecord.TelephoneNumber}");
                }
                else
                {
                    // Assign
                    await _graphPhone!.AssignNumberAsync(_currentlySelectedRecord.TelephoneNumber, _currentlySelectedRecord.PendingUserUpn);
                    AppendLog($"Phone Management: Assigned {_currentlySelectedRecord.TelephoneNumber} to {_currentlySelectedRecord.PendingUserUpn}");
                }

                // Update the record's assigned state
                _currentlySelectedRecord.AssignedUserUpn = _currentlySelectedRecord.PendingUserUpn;
            }

            // 2. Assign policies if Direct Routing + assigned
            if (_currentlySelectedRecord.CanAssignPolicies)
            {
                if (!string.IsNullOrWhiteSpace(_currentlySelectedRecord.PendingDialPlan) ||
                    !string.IsNullOrWhiteSpace(_currentlySelectedRecord.PendingVoiceRoutingPolicy))
                {
                    SetPhoneBusy(true, "Assigning policies...");

                    await _graphPhone!.AssignPoliciesAsync(
                        _currentlySelectedRecord.PendingUserUpn!,
                        _currentlySelectedRecord.PendingDialPlan,
                        _currentlySelectedRecord.PendingVoiceRoutingPolicy
                    );

                    AppendLog($"Phone Management: Assigned policies to {_currentlySelectedRecord.PendingUserUpn}");
                }
            }

            // 3. Clear pending changes (commit them)
            _currentlySelectedRecord.PendingUserUpn = _currentlySelectedRecord.AssignedUserUpn;
            _currentlySelectedRecord.PendingDialPlan = null;
            _currentlySelectedRecord.PendingVoiceRoutingPolicy = null;

            SidePanelStatusText.Text = "✓ Changes applied successfully.";
            SidePanelStatusText.Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(Microsoft.UI.Colors.Green);
            SidePanelStatusText.Visibility = Visibility.Visible;

            AppendLog($"Phone Management: Changes applied to {_currentlySelectedRecord.TelephoneNumber}");

            // Re-populate the side panel to reflect new state
            PopulateSidePanel(_currentlySelectedRecord);

            // Enable "Apply Changes" toolbar button if any rows have pending changes
            UpdateToolbarApplyButton();
        }
        catch (Exception ex)
        {
            AppendLog($"Phone Management: Apply failed — {ex.Message}");
            SidePanelStatusText.Text = $"✗ Error: {ex.Message}";
            SidePanelStatusText.Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(Microsoft.UI.Colors.Red);
            SidePanelStatusText.Visibility = Visibility.Visible;
        }
        finally
        {
            SetPhoneBusy(false);
            BtnApplySingleNumber.IsEnabled = _currentlySelectedRecord?.IsDirty ?? false;
        }
    }

    /// <summary>
    /// Update toolbar "Apply Changes" button based on whether any records have pending changes.
    /// </summary>
    private void UpdateToolbarApplyButton()
    {
        BtnApplyChanges.IsEnabled = _phoneRecords.Any(r => r.IsDirty);
    }

    private async void BtnApplyChanges_Click(object sender, RoutedEventArgs e)
    {
        var dirtyRecords = _phoneRecords.Where(r => r.IsDirty).ToList();
        if (dirtyRecords.Count == 0)
        {
            AppendLog("Phone Management: No pending changes to apply.");
            return;
        }

        BtnApplyChanges.IsEnabled = false;
        SetPhoneBusy(true, $"Applying changes to {dirtyRecords.Count} number(s)...");

        try
        {
            foreach (var record in dirtyRecords)
            {
                // Apply same logic as single number (you could refactor this into a shared method)
                if (record.PendingUserUpn != record.AssignedUserUpn)
                {
                    if (string.IsNullOrWhiteSpace(record.PendingUserUpn))
                    {
                        await _graphPhone!.UnassignNumberAsync(record.TelephoneNumber);
                    }
                    else
                    {
                        await _graphPhone!.AssignNumberAsync(record.TelephoneNumber, record.PendingUserUpn);
                    }
                    record.AssignedUserUpn = record.PendingUserUpn;
                }

                if (record.CanAssignPolicies &&
                    (!string.IsNullOrWhiteSpace(record.PendingDialPlan) ||
                     !string.IsNullOrWhiteSpace(record.PendingVoiceRoutingPolicy)))
                {
                    await _graphPhone!.AssignPoliciesAsync(
                        record.PendingUserUpn!,
                        record.PendingDialPlan,
                        record.PendingVoiceRoutingPolicy
                    );
                }

                record.PendingUserUpn = record.AssignedUserUpn;
                record.PendingDialPlan = null;
                record.PendingVoiceRoutingPolicy = null;
            }

            AppendLog($"Phone Management: Applied changes to {dirtyRecords.Count} number(s).");
        }
        catch (Exception ex)
        {
            AppendLog($"Phone Management: Batch apply failed — {ex.Message}");
        }
        finally
        {
            SetPhoneBusy(false);
            UpdateToolbarApplyButton();
            NumbersGrid.ItemsSource = null;  // Force refresh
            NumbersGrid.ItemsSource = _phoneRecords;
        }
    }
    
}