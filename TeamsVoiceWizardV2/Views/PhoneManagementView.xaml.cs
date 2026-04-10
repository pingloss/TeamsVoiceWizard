using CommunityToolkit.WinUI.UI.Controls;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using TeamsVoiceWizardV2.Models;
using TeamsVoiceWizardV2.ViewModels;

namespace TeamsVoiceWizardV2.Views;

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
        try
        {
            var dialogVm = _boundVm.CreateBulkImportViewModel();

            // On Uno Skia (macOS/Linux), this.XamlRoot may be null inside a TabView.
            // Fall back to the root XamlRoot from the main window content.
            var xamlRoot = XamlRoot ?? App.MainWindow?.Content?.XamlRoot;

            var dialog = new BulkImportDialog
            {
                XamlRoot   = xamlRoot,
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
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"BulkImport dialog failed: {ex}");
        }
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

    // ── DataGrid sort ─────────────────────────────────────────────────────────
    // Note: SortMemberPath is not available in the Uno DataGrid port.
    // Tag carries the sort property name instead; read it here.

    private void NumbersGrid_Sorting(object sender, DataGridColumnEventArgs e)
    {
        if (DataContext is not PhoneManagementViewModel vm) return;
        if (e.Column.Tag is not string sortPath) return;

        var newDirection = e.Column.SortDirection == DataGridSortDirection.Ascending
            ? DataGridSortDirection.Descending
            : DataGridSortDirection.Ascending;

        foreach (var col in NumbersGrid.Columns)
            col.SortDirection = null;
        e.Column.SortDirection = newDirection;

        var sorted = newDirection == DataGridSortDirection.Ascending
            ? vm.PhoneRecords.OrderBy(r => GetSortValue(r, sortPath)).ToList()
            : vm.PhoneRecords.OrderByDescending(r => GetSortValue(r, sortPath)).ToList();

        vm.PhoneRecords.Clear();
        foreach (var r in sorted)
            vm.PhoneRecords.Add(r);
    }

    private static string GetSortValue(PhoneNumberRecord r, string path) => path switch
    {
        nameof(PhoneNumberRecord.TelephoneNumber)         => r.TelephoneNumber ?? "",
        nameof(PhoneNumberRecord.NumberType)              => r.NumberType ?? "",
        nameof(PhoneNumberRecord.AssignedUserDisplayName) => r.AssignedUserDisplayName ?? "",
        nameof(PhoneNumberRecord.AssignmentStatus)        => r.AssignmentStatus ?? "",
        _ => ""
    };
}
