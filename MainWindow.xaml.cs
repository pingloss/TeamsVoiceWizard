using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using System.Diagnostics;
using System.Runtime.InteropServices;
using TeamsVoiceWizard.Services;
using TeamsVoiceWizard.ViewModels;
using Windows.Graphics;
using WinRT.Interop;

namespace TeamsVoiceWizard;

public sealed partial class MainWindow : Window
{
    private PowerShellHost? _ps;
    private readonly DispatcherQueue _dispatcher;

    private bool _startupRan;
    private bool _psReady;

    private ConfigurationViewModel? _configVm;

    public MainWindow()
    {
        InitializeComponent();

        Title       = "Teams Voice Wizard";
        _dispatcher = DispatcherQueue.GetForCurrentThread();

        MainTabView.Loaded += MainTabView_Loaded;
        Closed             += MainWindow_Closed;
    }

    private async void MainTabView_Loaded(object sender, Microsoft.UI.Xaml.RoutedEventArgs e)
    {
        if (_startupRan) return;
        _startupRan = true;

        TrySetInitialWindowSize(width: 1140, height: 880);

        // Wire ViewModels first so AppendLog is available before PS init starts.
        WireViewModels();

        _configVm!.AppendLog("Teams Voice Wizard started.");

        var sw = Stopwatch.StartNew();
        _configVm.AppendLog($"[Perf] Startup: begin @ {sw.ElapsedMilliseconds}ms");

        await InitPowerShellAsync();

        _configVm.AppendLog($"[Perf] Startup: PS init complete @ {sw.ElapsedMilliseconds}ms");
    }

    private void WireViewModels()
    {
        // ── Configuration ViewModel ───────────────────────────────────────────
        var configServices = new ConfigurationHostServices(
            IsPowerShellReady:            () => _psReady,
            ConnectGraphAsync:            scopes => _ps!.ConnectGraphAsync(scopes),
            ConnectTeamsAsync:            () => _ps!.ConnectTeamsAsync(),
            IsGraphConnectedAsync:        () => _ps!.IsGraphConnectedAsync(),
            IsTeamsConnectedAsync:        () => _ps!.IsTeamsConnectedAsync(),
            CreateDomainsAndGetTxtAsync:  (ep, site, country) => _ps!.CreateDomainsAndGetTxtAsync(ep, site, country),
            VerifyDomainsAsync:           () => _ps!.VerifyDomainsAsync(),
            LoadLicenseInventoryAsync:    minFree => _ps!.LoadLicenseInventoryAsync(minFree),
            CreateTestObjectsAsync:       (sku, country) => _ps!.CreateTestObjectsAsync(sku, country),
            CleanupTestObjectsAsync:      (users, ras) => _ps!.CleanupTestObjectsAsync(users, ras),
            ApplyVoiceConfigurationAsync: (site, country, derived) => _ps!.ApplyVoiceConfigurationAsync(site, country, derived)
        );

        _configVm = new ConfigurationViewModel(configServices);
        ConfigurationHost.DataContext = _configVm;

        // Enable the Phone Management tab whenever Graph connects.
        _configVm.PropertyChanged += (_, args) =>
        {
            if (args.PropertyName == nameof(ConfigurationViewModel.GraphConnected))
                PhoneTab.IsEnabled = _configVm.GraphConnected;
        };

        // ── Phone Management ViewModel ────────────────────────────────────────
        // Shares _configVm.AppendLog so phone management messages appear in the same log box.
        var phoneServices = new PhoneManagementHostServices(
            IsGraphConnected:             () => _configVm.GraphConnected,
            TryEnsurePowerShell:          EnsurePowerShellReady,
            GetGraphAccessTokenAsync:     () => _ps!.GetGraphAccessTokenAsync(),
            GetDialPlansAsync:            () => _ps!.GetDialPlansAsync(),
            GetVoiceRoutingPoliciesAsync: () => _ps!.GetVoiceRoutingPoliciesAsync(),
            Log:                          line => _configVm.AppendLog(line)
        );

        PhoneManagementHost.DataContext = new PhoneManagementViewModel(phoneServices);
    }

    private void MainWindow_Closed(object sender, WindowEventArgs args)
    {
        try { _ps?.Dispose(); }
        catch (Exception ex) { Debug.WriteLine($"PowerShell cleanup failed: {ex}"); }
        finally
        {
            _ps      = null;
            _psReady = false;
        }
    }

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    /// <summary>
    /// Sets the initial window size in device-independent units (DIPs).
    /// AppWindow.Resize takes physical pixels, so we scale by the window's DPI to
    /// keep the window the same effective size on Hi-DPI displays.
    /// </summary>
    private void TrySetInitialWindowSize(int width, int height)
    {
        try
        {
            var hwnd      = WindowNative.GetWindowHandle(this);
            var windowId  = Win32Interop.GetWindowIdFromWindow(hwnd);
            var appWindow = AppWindow.GetFromWindowId(windowId);

            double scale = 1.0;
            try
            {
                var dpi = GetDpiForWindow(hwnd);
                if (dpi > 0) scale = dpi / 96.0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetDpiForWindow failed, defaulting to 1.0: {ex}");
            }

            appWindow.Resize(new SizeInt32(
                (int)Math.Round(width  * scale),
                (int)Math.Round(height * scale)));
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
                _configVm!.AppendLog($"[Startup] PowerShell module not found: {modulePath}");
                _configVm.AppendLog("[Startup] Ensure core\\TeamsVoiceWizard.Core.psm1 is copied to the output folder.");
                _psReady = false;
                _configVm.SetPowerShellReady(false);
                return;
            }

            _configVm!.AppendLog("[Startup] Initialising PowerShell runspace...");

            _ps = await Task.Run(() =>
                new PowerShellHost(modulePath,
                    line => _dispatcher.TryEnqueue(() => _configVm!.AppendLog(line)),
                    diagnostics: false)
            );

            _psReady = true;
            _configVm.SetPowerShellReady(true);

            // Validate connections and propagate initial state to the ViewModel.
            var graphConnected = await _ps.IsGraphConnectedAsync();
            var teamsConnected = await _ps.IsTeamsConnectedAsync();

            _configVm.GraphConnected = graphConnected;
            _configVm.TeamsConnected = teamsConnected;

            _configVm.AppendLog("[Startup] PowerShell runspace ready.");
        }
        catch (Exception ex)
        {
            _psReady = false;
            _ps      = null;
            _configVm!.SetPowerShellReady(false);
            _configVm.AppendLog($"[Startup] PowerShell init failed: {ex.Message}");
        }
    }

    /// <summary>Used by PhoneManagementHostServices to guard PS calls.</summary>
    private bool EnsurePowerShellReady(bool logIfNotReady = true)
    {
        if (_psReady && _ps is not null) return true;

        if (logIfNotReady)
            _configVm?.AppendLog("PowerShell host not ready (startup failed or module missing).");

        return false;
    }
}
