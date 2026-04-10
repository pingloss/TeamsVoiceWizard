using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using System.Diagnostics;
using TeamsVoiceWizardV2.Services;
using TeamsVoiceWizardV2.ViewModels;

#if WINDOWS
using Microsoft.UI.Windowing;
using Microsoft.UI;
using System.Runtime.InteropServices;
using Windows.Graphics;
using WinRT.Interop;
#endif

namespace TeamsVoiceWizardV2;

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

        // On Uno Skia, TabView passes infinite height to its content so star-rows
        // collapse. Drive content height from the window size instead.
        SizeChanged += OnWindowSizeChanged;
    }

    private void OnTabViewFirstLayout(object? sender, object e)
    {
        if (MainTabView.ActualHeight <= 0) return;
        MainTabView.LayoutUpdated -= OnTabViewFirstLayout; // fire once
        ApplyTabContentHeight(MainTabView.ActualHeight);
    }

    private void OnWindowSizeChanged(object sender, WindowSizeChangedEventArgs args)
    {
        ApplyTabContentHeight(args.Size.Height);
    }

    private void ApplyTabContentHeight(double windowHeight)
    {
        // Measure the actual tab strip height once it has laid out; fall back to 48.
        var tabStripHeight = MainTabView.ActualHeight > 0 && ConfigurationHost.ActualHeight > 0
            ? MainTabView.ActualHeight - ConfigurationHost.ActualHeight
            : 48;

        var contentHeight = Math.Max(200, windowHeight - tabStripHeight);
        ConfigurationHost.Height   = contentHeight;
        PhoneManagementHost.Height = contentHeight;
    }

    private async void MainTabView_Loaded(object sender, RoutedEventArgs e)
    {
        if (_startupRan) return;
        _startupRan = true;

        TrySetInitialWindowSize(width: 1150, height: 850);

        // Wire ViewModels first so AppendLog is available before PS init starts.
        WireViewModels();

        // Apply initial content height after layout (window size known by this point).
        MainTabView.LayoutUpdated += OnTabViewFirstLayout;

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

#if WINDOWS
    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);
#endif

    /// <summary>
    /// Sets the initial window size in device-independent units (DIPs).
    /// On Windows, AppWindow.Resize takes physical pixels so we scale by DPI.
    /// On Skia/desktop Uno handles window sizing natively — no action needed.
    /// </summary>
    private void TrySetInitialWindowSize(int width, int height)
    {
#if WINDOWS
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
#else
        try
        {
            var view = Windows.UI.ViewManagement.ApplicationView.GetForCurrentView();
            view.SetPreferredMinSize(new Windows.Foundation.Size(width, height));
            view.TryResizeView(new Windows.Foundation.Size(width, height));
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"TrySetInitialWindowSize (Skia) failed: {ex}");
        }
#endif
    }

    private async Task InitPowerShellAsync()
    {
        try
        {
            var modulePath = Path.Combine(AppContext.BaseDirectory, "core", "TeamsVoiceWizard.Core.psm1");

            if (!File.Exists(modulePath))
            {
                _configVm!.AppendLog($"[Startup] PowerShell module not found: {modulePath}");
                _configVm.AppendLog("[Startup] Ensure core/TeamsVoiceWizard.Core.psm1 is copied to the output folder.");
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
