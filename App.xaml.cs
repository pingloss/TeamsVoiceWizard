using Microsoft.UI.Xaml;

namespace TeamsVoiceWizard;

    public partial class App : Application
{
    private Window? _window;

    public App()
    {
        InitializeComponent();
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        _window = new MainWindow();
        _window.Activate();
    }

    /// <summary>Returns the main window so dialogs can get its HWND for file pickers.</summary>
    public Window GetMainWindow() => _window!;
}