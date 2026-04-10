using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.Storage.Pickers;
using TeamsVoiceWizardV2.ViewModels;

#if WINDOWS
using WinRT.Interop;
#endif

namespace TeamsVoiceWizardV2.Views;

public sealed partial class BulkImportDialog : ContentDialog
{
    private BulkImportViewModel? _vm;

    public BulkImportDialog()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
        PrimaryButtonClick += OnPrimaryButtonClick;
        Loaded             += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        // Wire EmptyPlaceholder visibility — inverse of HasRows
        if (_vm is not null)
            UpdatePlaceholderVisibility(_vm.HasRows);
    }

    private void OnDataContextChanged(FrameworkElement sender, DataContextChangedEventArgs args)
    {
        if (_vm is not null)
        {
            _vm.PropertyChanged   -= OnVmPropertyChanged;
            _vm.ApplyCompleted    -= OnApplyCompleted;
        }

        _vm = args.NewValue as BulkImportViewModel;

        if (_vm is not null)
        {
            _vm.PropertyChanged   += OnVmPropertyChanged;
            _vm.ApplyCompleted    += OnApplyCompleted;
            IsPrimaryButtonEnabled = _vm.CanApply;
        }
    }

    private void OnVmPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (_vm is null) return;

        if (e.PropertyName is nameof(BulkImportViewModel.CanApply))
            IsPrimaryButtonEnabled = _vm.CanApply;

        if (e.PropertyName is nameof(BulkImportViewModel.HasRows))
            UpdatePlaceholderVisibility(_vm.HasRows);
    }

    private void OnApplyCompleted(object? sender, EventArgs e)
    {
        // After apply finishes, keep Primary button disabled — user reads results then clicks Close.
        IsPrimaryButtonEnabled = false;
    }

    private void UpdatePlaceholderVisibility(bool hasRows)
    {
        EmptyPlaceholder.Visibility = hasRows ? Visibility.Collapsed : Visibility.Visible;
    }

    // ── Primary button = Apply ────────────────────────────────────────────────

    private async void OnPrimaryButtonClick(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        // Defer close so we can run the apply command first.
        var deferral = args.GetDeferral();

        if (_vm?.ApplyCommand.CanExecute(null) == true)
            await _vm.ApplyCommand.ExecuteAsync(null);

        // Don't close the dialog automatically — user reads results then clicks Close.
        args.Cancel = true;
        deferral.Complete();
    }

    // ── File picker ───────────────────────────────────────────────────────────

    private async void BrowseButton_Click(object sender, RoutedEventArgs e)
    {
        if (_vm is null) return;

        var picker = new FileOpenPicker
        {
            SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
            FileTypeFilter         = { ".csv", ".xlsx" }
        };

#if WINDOWS
        // Required on Windows to associate the picker with the app window
        var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(App.MainWindow);
        WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);
#endif

        var file = await picker.PickSingleFileAsync();
        if (file is null) return;

        var stream = await file.OpenStreamForReadAsync();
        await _vm.LoadFileAsync(stream, file.Name, file.FileType);
    }
}
