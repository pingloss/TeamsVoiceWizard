using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using TeamsVoiceWizard.ViewModels;

namespace TeamsVoiceWizard.Views;

public sealed partial class ConfigurationView : UserControl
{
    private ConfigurationViewModel? _vm;

    public ConfigurationView()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
        Unloaded           += OnUnloaded;
    }

    private void OnDataContextChanged(FrameworkElement sender, DataContextChangedEventArgs args)
    {
        if (_vm is not null)
        {
            _vm.LogScrollRequested    -= OnLogScrollRequested;
            _vm.OutputScrollRequested -= OnOutputScrollRequested;
        }

        _vm = args.NewValue as ConfigurationViewModel;

        if (_vm is not null)
        {
            _vm.LogScrollRequested    += OnLogScrollRequested;
            _vm.OutputScrollRequested += OnOutputScrollRequested;

            // Sync TextBoxes with ViewModel's initial values.
            // Plain {Binding} on TextBox.Text is OneWay by default so we prime manually.
            GammaBox.Text   = _vm.GammaEndpoint;
            SiteBox.Text    = _vm.SiteName;
            CountryBox.Text = _vm.Country;
        }
    }

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        if (_vm is not null)
        {
            _vm.LogScrollRequested    -= OnLogScrollRequested;
            _vm.OutputScrollRequested -= OnOutputScrollRequested;
            _vm = null;
        }
    }

    // ── Scroll-to-bottom event handlers ──────────────────────────────────────

    private void OnLogScrollRequested(object? sender, EventArgs e)
    {
        try { LogScroll?.ChangeView(null, double.MaxValue, null); }
        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"LogScroll ChangeView failed: {ex}"); }
    }

    private void OnOutputScrollRequested(object? sender, EventArgs e)
    {
        try { OutputScroll?.ChangeView(null, double.MaxValue, null); }
        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"OutputScroll ChangeView failed: {ex}"); }
    }

    // ── TextBox relay handlers ────────────────────────────────────────────────
    // Classic {Binding} does not guarantee source updates on every keystroke for
    // TextBox.Text, so we relay through code-behind to keep guardrails reactive.

    private void GammaBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_vm is not null)
            _vm.GammaEndpoint = ((TextBox)sender).Text;
    }

    private void SiteBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_vm is not null)
            _vm.SiteName = ((TextBox)sender).Text;
    }

    private void CountryBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_vm is not null)
            _vm.Country = ((TextBox)sender).Text.Trim().ToUpperInvariant();
    }
}
