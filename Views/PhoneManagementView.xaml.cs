using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Data;
using TeamsVoiceWizard.ViewModels;

namespace TeamsVoiceWizard.Views;

public sealed partial class PhoneManagementView : UserControl
{
    private PhoneManagementViewModel? _boundVm;

    public PhoneManagementView()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
        Unloaded += OnUnloaded;
    }

    private void OnDataContextChanged(FrameworkElement sender, DataContextChangedEventArgs args)
    {
        if (_boundVm is not null)
        {
            _boundVm.GridRefreshRequested -= OnGridRefreshRequested;
            _boundVm = null;
        }

        if (args.NewValue is PhoneManagementViewModel vm)
        {
            _boundVm = vm;
            vm.GridRefreshRequested += OnGridRefreshRequested;
        }
    }

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        if (_boundVm is not null)
        {
            _boundVm.GridRefreshRequested -= OnGridRefreshRequested;
            _boundVm = null;
        }
    }

    private void OnGridRefreshRequested(object? sender, EventArgs e)
    {
        var src = NumbersGrid.ItemsSource;
        NumbersGrid.ItemsSource = null;
        NumbersGrid.ItemsSource = src;
    }

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
}
