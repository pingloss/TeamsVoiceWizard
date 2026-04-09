using CommunityToolkit.WinUI.UI.Controls;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using TeamsVoiceWizard.ViewModels;

namespace TeamsVoiceWizard.Views;

public sealed partial class PhoneManagementView : UserControl
{
    private PhoneManagementViewModel? _boundVm;

    public PhoneManagementView()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
        Unloaded           += OnUnloaded;
    }

    private void OnDataContextChanged(FrameworkElement sender, DataContextChangedEventArgs args)
    {
        if (_boundVm is not null)
        {
            _boundVm.GridRefreshRequested -= OnGridRefreshRequested;
            _boundVm.BulkImportRequested  -= OnBulkImportRequested;
            _boundVm = null;
        }

        if (args.NewValue is PhoneManagementViewModel vm)
        {
            _boundVm = vm;
            vm.GridRefreshRequested += OnGridRefreshRequested;
            vm.BulkImportRequested  += OnBulkImportRequested;
        }
    }

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        if (_boundVm is not null)
        {
            _boundVm.GridRefreshRequested -= OnGridRefreshRequested;
            _boundVm.BulkImportRequested  -= OnBulkImportRequested;
            _boundVm = null;
        }
    }

    private void OnGridRefreshRequested(object? sender, EventArgs e)
    {
        var src = NumbersGrid.ItemsSource;
        NumbersGrid.ItemsSource = null;
        NumbersGrid.ItemsSource = src;
    }

    // ── Bulk Import dialog ────────────────────────────────────────────────────

    private async void OnBulkImportRequested(object? sender, EventArgs e)
    {
        if (_boundVm is null) return;

        var dialogVm = _boundVm.CreateBulkImportViewModel();

        var dialog = new BulkImportDialog
        {
            XamlRoot   = XamlRoot,
            DataContext = dialogVm
        };

        // Reload numbers after a successful apply so the grid reflects the changes
        dialogVm.ApplyCompleted += async (_, _) =>
        {
            if (_boundVm?.LoadNumbersCommand.CanExecute(null) == true)
                await _boundVm.LoadNumbersCommand.ExecuteAsync(null);
        };

        await dialog.ShowAsync();
    }

    // ── ComboBox event handlers ───────────────────────────────────────────────

    private async void UserComboBox_DropDownOpened(object sender, object e)
    {
        if (DataContext is PhoneManagementViewModel vm)
            await vm.OnUserComboDropDownOpenedAsync().ConfigureAwait(true);
    }

    private async void DialPlanComboBox_DropDownOpened(object sender, object e)
    {
        if (DataContext is PhoneManagementViewModel vm)
            await vm.OnDialPlanComboDropDownOpenedAsync().ConfigureAwait(true);
    }

    private async void VoiceRoutingPolicyComboBox_DropDownOpened(object sender, object e)
    {
        if (DataContext is PhoneManagementViewModel vm)
            await vm.OnVoiceRoutingComboDropDownOpenedAsync().ConfigureAwait(true);
    }

    private void NumbersGrid_Sorting(object sender, DataGridColumnEventArgs e)
    {
        if (DataContext is not PhoneManagementViewModel vm) return;
        if (e.Column.Tag is not string sortPath) return;

        var newDirection = vm.SortRecords(sortPath);

        foreach (var col in NumbersGrid.Columns)
            col.SortDirection = null;
        e.Column.SortDirection = newDirection;
    }
}
