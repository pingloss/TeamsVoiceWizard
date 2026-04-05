using System.Collections.ObjectModel;
using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.UI;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Media;
using TeamsVoiceWizard.Models;
using TeamsVoiceWizard.Services;

namespace TeamsVoiceWizard.ViewModels;

public partial class PhoneManagementViewModel : ObservableObject
{
    private readonly PhoneManagementHostServices _host;
    private readonly DispatcherQueue _dispatcher;

    private GraphPhoneService? _graphPhone;
    private readonly PolicyCaches _policyCaches = new();
    private readonly Dictionary<string, string> _graphPolicyIdCache = new();
    private readonly Dictionary<string, string> _usersCache = new();

    private bool _isLoadingUsers;
    private bool _isLoadingPolicies;
    private bool _suppressPolicyCombo;
    private bool _suppressUserCombo;
    private PhoneNumberRecord? _recordSubscribed;

    public PhoneManagementViewModel(PhoneManagementHostServices host)
    {
        _host = host;
        _dispatcher = DispatcherQueue.GetForCurrentThread();
        PhoneRecords.CollectionChanged += (_, e) =>
        {
            if (e.NewItems is not null)
            {
                foreach (PhoneNumberRecord r in e.NewItems)
                    r.PropertyChanged += OnPhoneRecordPropertyChanged;
            }

            if (e.OldItems is not null)
            {
                foreach (PhoneNumberRecord r in e.OldItems)
                    r.PropertyChanged -= OnPhoneRecordPropertyChanged;
            }
        };
    }

    public ObservableCollection<PhoneNumberRecord> PhoneRecords { get; } = new();

    public ObservableCollection<ComboPickItem> DialPlanOptions { get; } = new();
    public ObservableCollection<ComboPickItem> VoiceRoutingOptions { get; } = new();
    public ObservableCollection<ComboPickItem> UserOptions { get; } = new();

    public event EventHandler? GridRefreshRequested;

    [ObservableProperty]
    public partial PhoneNumberRecord? SelectedRecord { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsLoadNumbersEnabled))]
    public partial bool IsPhoneBusy { get; set; }

    [ObservableProperty]
    public partial string PhoneStatusMessage { get; set; } = "";

    [ObservableProperty]
    public partial string DetailPhoneNumber { get; set; } = "";

    [ObservableProperty]
    public partial string DetailNumberType { get; set; } = "";

    [ObservableProperty]
    public partial string DetailStatus { get; set; } = "";

    [ObservableProperty]
    public partial string DetailCurrentUser { get; set; } = "";

    [ObservableProperty]
    public partial bool UserLoadingRingActive { get; set; }

    [ObservableProperty]
    public partial bool DialPlanLoadingRingActive { get; set; }

    [ObservableProperty]
    public partial bool VRPolicyLoadingRingActive { get; set; }

    [ObservableProperty]
    public partial ComboPickItem? SelectedUserItem { get; set; }

    [ObservableProperty]
    public partial ComboPickItem? SelectedDialPlanItem { get; set; }

    [ObservableProperty]
    public partial ComboPickItem? SelectedVoiceRoutingItem { get; set; }

    [ObservableProperty]
    public partial string SidePanelStatusText { get; set; } = "";

    [ObservableProperty]
    public partial bool SidePanelStatusVisible { get; set; }

    [ObservableProperty]
    public partial Brush SidePanelStatusBrush { get; set; } = new SolidColorBrush(Colors.Gray);

    public bool IsSidePanelVisible => SelectedRecord is not null;

    public bool IsPolicySectionVisible => SelectedRecord?.CanAssignPolicies == true;

    public bool IsLoadNumbersEnabled => !IsPhoneBusy;

    partial void OnSelectedRecordChanged(PhoneNumberRecord? oldValue, PhoneNumberRecord? newValue)
    {
        if (_recordSubscribed is not null)
            _recordSubscribed.PropertyChanged -= OnSelectedRecordIsDirtyChanged;
        _recordSubscribed = newValue;
        if (_recordSubscribed is not null)
            _recordSubscribed.PropertyChanged += OnSelectedRecordIsDirtyChanged;

        OnPropertyChanged(nameof(IsSidePanelVisible));
        OnPropertyChanged(nameof(IsPolicySectionVisible));
        SyncDetailFromRecord();
        ApplySingleChangeCommand.NotifyCanExecuteChanged();

        if (newValue is null)
            return;

        _ = LoadSidePanelForRecordAsync(newValue);
    }

    private void OnSelectedRecordIsDirtyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(PhoneNumberRecord.IsDirty))
            ApplySingleChangeCommand.NotifyCanExecuteChanged();
    }

    private void OnPhoneRecordPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(PhoneNumberRecord.IsDirty))
            ApplyBatchChangesCommand.NotifyCanExecuteChanged();
    }

    partial void OnSelectedUserItemChanged(ComboPickItem? oldValue, ComboPickItem? newValue)
    {
        if (_suppressUserCombo || SelectedRecord is null) return;
        SelectedRecord.PendingUserUpn = newValue?.Value;
        ApplySingleChangeCommand.NotifyCanExecuteChanged();
    }

    partial void OnSelectedDialPlanItemChanged(ComboPickItem? oldValue, ComboPickItem? newValue)
    {
        if (_suppressPolicyCombo || SelectedRecord is null) return;
        SelectedRecord.PendingDialPlan = newValue?.Value;
        ApplySingleChangeCommand.NotifyCanExecuteChanged();
    }

    partial void OnSelectedVoiceRoutingItemChanged(ComboPickItem? oldValue, ComboPickItem? newValue)
    {
        if (_suppressPolicyCombo || SelectedRecord is null) return;
        SelectedRecord.PendingVoiceRoutingPolicy = newValue?.Value;
        ApplySingleChangeCommand.NotifyCanExecuteChanged();
    }

    private void SyncDetailFromRecord()
    {
        if (SelectedRecord is null)
        {
            DetailPhoneNumber = "";
            DetailNumberType = "";
            DetailStatus = "";
            DetailCurrentUser = "";
            return;
        }

        DetailPhoneNumber = SelectedRecord.TelephoneNumber ?? "(unknown)";
        DetailNumberType = SelectedRecord.NumberType ?? "(unknown)";
        DetailStatus = SelectedRecord.AssignmentStatus ?? "(unknown)";
        DetailCurrentUser = SelectedRecord.AssignedUserDisplayName ?? "(unassigned)";
    }

    private GraphPhoneService RequireGraphPhone() =>
        _graphPhone ??= new GraphPhoneService(_host.GetGraphAccessTokenAsync);

    private void RunOnUi(Action action)
    {
        if (_dispatcher.HasThreadAccess)
            action();
        else
            _dispatcher.TryEnqueue(() => action());
    }

    [RelayCommand]
    private async Task LoadNumbersAsync()
    {
        if (!_host.IsGraphConnected())
        {
            _host.Log("Phone Management: Graph not connected.");
            return;
        }

        IsPhoneBusy = true;
        PhoneStatusMessage = "Loading phone numbers...";
        try
        {
            var graph = RequireGraphPhone();

            var records = await graph.GetNumberAssignmentsAsync().ConfigureAwait(false);

            var assignedIds = records
                .Where(r => !string.IsNullOrWhiteSpace(r.AssignmentTargetId))
                .Select(r => r.AssignmentTargetId!)
                .Distinct()
                .ToList();

            RunOnUi(() =>
            {
                PhoneStatusMessage = $"Resolving {assignedIds.Count} user(s)...";
            });

            var resolved = assignedIds.Count > 0
                ? await graph.ResolveUsersAsync(assignedIds).ConfigureAwait(false)
                : new Dictionary<string, (string DisplayName, string Upn)>();

            if (assignedIds.Count > 0)
            {
                RunOnUi(() => { PhoneStatusMessage = "Building policy cache..."; });

                var configTasks = assignedIds
                    .Select(id => graph.GetUserTeamsConfigurationAsync(id, log: null))
                    .ToList();

                var configs = await Task.WhenAll(configTasks).ConfigureAwait(false);

                RunOnUi(() =>
                {
                    foreach (var config in configs)
                    {
                        foreach (var (policyType, (displayName, policyId)) in config)
                            _graphPolicyIdCache[$"{policyType}:{displayName}"] = policyId;
                    }

                    _host.Log($"Phone Management: Cached {_graphPolicyIdCache.Count} Graph policy ID(s).");
                });
            }

            RunOnUi(() =>
            {
                PhoneRecords.Clear();
                foreach (var r in records)
                {
                    if (r.AssignmentTargetId is not null &&
                        resolved.TryGetValue(r.AssignmentTargetId, out var user))
                    {
                        r.AssignedUserDisplayName = user.DisplayName;
                        r.AssignedUserUpn = user.Upn;
                    }

                    PhoneRecords.Add(r);
                }

                _host.Log($"Phone Management: Loaded {PhoneRecords.Count} number(s).");
                ApplyBatchChangesCommand.NotifyCanExecuteChanged();
            });
        }
        catch (Exception ex)
        {
            RunOnUi(() => _host.Log($"Phone Management: Load failed — {ex.Message}"));
        }
        finally
        {
            RunOnUi(() =>
            {
                IsPhoneBusy = false;
                PhoneStatusMessage = "";
            });
        }
    }

    private async Task LoadSidePanelForRecordAsync(PhoneNumberRecord record)
    {
        try
        {
            await EnsureUsersLoadedAsync(record).ConfigureAwait(false);

            if (record.CanAssignPolicies)
            {
                RunOnUi(() => PopulatePolicyComboBoxes(record));
                await EnsurePoliciesLoadedAsync(record).ConfigureAwait(false);
            }

            RunOnUi(() => ApplySingleChangeCommand.NotifyCanExecuteChanged());
        }
        catch (Exception ex)
        {
            _host.Log($"Phone Management: Side panel load error — {ex.Message}");
        }
    }

    public Task OnUserComboDropDownOpenedAsync() => EnsureUsersLoadedAsync(SelectedRecord);

    public Task OnDialPlanComboDropDownOpenedAsync() => EnsurePoliciesLoadedIfNeededAsync();

    public Task OnVoiceRoutingComboDropDownOpenedAsync() => EnsurePoliciesLoadedIfNeededAsync();

    private async Task EnsurePoliciesLoadedIfNeededAsync()
    {
        if (_isLoadingPolicies || (_policyCaches.DialPlans.Count > 0 && _policyCaches.VoiceRoutingPolicies.Count > 0))
            return;
        await EnsurePoliciesLoadedAsync(SelectedRecord).ConfigureAwait(true);
    }

    private async Task EnsureUsersLoadedAsync(PhoneNumberRecord? record = null)
    {
        if (_usersCache.Count > 0)
        {
            if (record is not null && !string.IsNullOrWhiteSpace(record.AssignmentTargetId))
                SelectUserByObjectId(record.AssignmentTargetId);
            return;
        }

        if (_isLoadingUsers) return;

        _isLoadingUsers = true;
        RunOnUi(() => UserLoadingRingActive = true);

        try
        {
            var graph = RequireGraphPhone();
            var licensedUsers = await graph.GetTeamsPhoneLicensedUsersAsync().ConfigureAwait(false);

            RunOnUi(() =>
            {
                UserOptions.Clear();
                foreach (var (userId, _, displayName) in licensedUsers)
                {
                    _usersCache[userId] = displayName;
                    UserOptions.Add(new ComboPickItem(displayName, userId));
                }

                _host.Log($"Phone Management: Loaded {_usersCache.Count} licensed Teams Phone user(s).");

                if (record is not null && !string.IsNullOrWhiteSpace(record.AssignmentTargetId))
                    SelectUserByObjectId(record.AssignmentTargetId);
            });
        }
        catch (Exception ex)
        {
            RunOnUi(() => _host.Log($"Phone Management: Failed to load users — {ex.Message}"));
        }
        finally
        {
            _isLoadingUsers = false;
            RunOnUi(() => UserLoadingRingActive = false);
        }
    }

    private async Task EnsurePoliciesLoadedAsync(PhoneNumberRecord? record = null)
    {
        if (_policyCaches.DialPlans.Count > 0 && _policyCaches.VoiceRoutingPolicies.Count > 0)
        {
            if (record is not null)
                await PreSelectCurrentValuesAsync(record).ConfigureAwait(true);
            return;
        }

        if (!_host.TryEnsurePowerShell(false))
            return;

        _isLoadingPolicies = true;
        RunOnUi(() =>
        {
            DialPlanLoadingRingActive = true;
            VRPolicyLoadingRingActive = true;
        });

        try
        {
            var dialTask = _policyCaches.DialPlans.Count == 0
                ? _host.GetDialPlansAsync()
                : Task.FromResult(_policyCaches.DialPlans.ToList());

            var vrTask = _policyCaches.VoiceRoutingPolicies.Count == 0
                ? _host.GetVoiceRoutingPoliciesAsync()
                : Task.FromResult(_policyCaches.VoiceRoutingPolicies.ToList());

            await Task.WhenAll(dialTask, vrTask).ConfigureAwait(false);

            var dialPlans = await dialTask.ConfigureAwait(false);
            var vrPolicies = await vrTask.ConfigureAwait(false);

            RunOnUi(() =>
            {
                if (dialPlans?.Count > 0)
                {
                    _policyCaches.DialPlans = dialPlans;
                    _host.Log($"Phone Management: Loaded {dialPlans.Count} dial plan(s).");
                }

                if (vrPolicies?.Count > 0)
                {
                    _policyCaches.VoiceRoutingPolicies = vrPolicies;
                    _host.Log($"Phone Management: Loaded {vrPolicies.Count} voice routing policy(s).");
                }

                if (record is not null && record.CanAssignPolicies)
                    PopulatePolicyComboBoxes(record);
            });

            if (record is not null)
                await PreSelectCurrentValuesAsync(record).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            RunOnUi(() => _host.Log($"Phone Management: Failed to load policies — {ex.Message}"));
        }
        finally
        {
            _isLoadingPolicies = false;
            RunOnUi(() =>
            {
                DialPlanLoadingRingActive = false;
                VRPolicyLoadingRingActive = false;
            });
        }
    }

    private async Task PreSelectCurrentValuesAsync(PhoneNumberRecord record)
    {
        if (!record.CanAssignPolicies) return;

        if (!record.PoliciesFetched && !string.IsNullOrWhiteSpace(record.AssignmentTargetId))
        {
            try
            {
                var graph = RequireGraphPhone();
                var allPolicies = await graph
                    .GetUserTeamsConfigurationAsync(record.AssignmentTargetId, _host.Log)
                    .ConfigureAwait(false);

                RunOnUi(() =>
                {
                    foreach (var (policyType, (displayName, policyId)) in allPolicies)
                        _graphPolicyIdCache[$"{policyType}:{displayName}"] = policyId;

                    record.CurrentDialPlan = allPolicies.TryGetValue("TenantDialPlan", out var dp)
                        ? dp.DisplayName
                        : null;
                    record.CurrentVoiceRoutingPolicy = allPolicies.TryGetValue("OnlineVoiceRoutingPolicy", out var vrp)
                        ? vrp.DisplayName
                        : null;

                    _host.Log($"[PolicyFetch] {record.AssignedUserUpn} — " +
                              $"DialPlan='{record.CurrentDialPlan ?? "none"}' " +
                              $"VRP='{record.CurrentVoiceRoutingPolicy ?? "none"}'");
                });
            }
            catch (Exception ex)
            {
                RunOnUi(() =>
                    _host.Log($"Phone Management: Could not fetch current policies — {ex.Message}"));
            }
            finally
            {
                record.PoliciesFetched = true;
            }
        }

        RunOnUi(() =>
        {
            SelectDialPlanByMatch(record.CurrentDialPlan);
            SelectVoiceRoutingByMatch(record.CurrentVoiceRoutingPolicy);
        });
    }

    private void PopulatePolicyComboBoxes(PhoneNumberRecord record)
    {
        DialPlanOptions.Clear();
        VoiceRoutingOptions.Clear();
        DialPlanOptions.Add(new ComboPickItem("(None - Keep Current)", null));
        VoiceRoutingOptions.Add(new ComboPickItem("(None - Keep Current)", null));

        foreach (var plan in _policyCaches.DialPlans)
            DialPlanOptions.Add(new ComboPickItem(plan.DisplayName, plan.Id));

        foreach (var policy in _policyCaches.VoiceRoutingPolicies)
            VoiceRoutingOptions.Add(new ComboPickItem(policy.DisplayName, policy.Id));
    }

    private string? ResolveGraphPolicyId(string policyType, string? psIdentity)
    {
        if (string.IsNullOrWhiteSpace(psIdentity)) return null;

        var name = psIdentity.StartsWith("Tag:", StringComparison.OrdinalIgnoreCase)
            ? psIdentity[4..]
            : psIdentity;

        var key = $"{policyType}:{name}";
        if (_graphPolicyIdCache.TryGetValue(key, out var graphId))
            return graphId;

        _host.Log($"Phone Management: Graph policy ID not cached for '{key}'. " +
                  "Click 'Load Numbers' to rebuild the cache.");
        return null;
    }

    private void SelectUserByObjectId(string objectId)
    {
        _suppressUserCombo = true;
        try
        {
            var match = UserOptions.FirstOrDefault(o =>
                string.Equals(o.Value, objectId, StringComparison.OrdinalIgnoreCase));
            SelectedUserItem = match;
        }
        finally
        {
            _suppressUserCombo = false;
        }
    }

    private static string? StripTagPrefix(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return value;
        return value.StartsWith("Tag:", StringComparison.OrdinalIgnoreCase) ? value[4..] : value;
    }

    private void SelectDialPlanByMatch(string? tagValue)
    {
        SelectPolicyPick(DialPlanOptions, tagValue, item => SelectedDialPlanItem = item);
    }

    private void SelectVoiceRoutingByMatch(string? tagValue)
    {
        SelectPolicyPick(VoiceRoutingOptions, tagValue, item => SelectedVoiceRoutingItem = item);
    }

    private void SelectPolicyPick(
        ObservableCollection<ComboPickItem> options,
        string? tagValue,
        Action<ComboPickItem?> assign)
    {
        _suppressPolicyCombo = true;
        try
        {
            if (options.Count == 0)
            {
                assign(null);
                return;
            }

            if (string.IsNullOrWhiteSpace(tagValue))
            {
                assign(options[0]);
                return;
            }

            foreach (var o in options)
            {
                var itemVal = StripTagPrefix(o.Value) ?? "";

                if (string.Equals(itemVal, tagValue, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(o.Label, tagValue, StringComparison.OrdinalIgnoreCase))
                {
                    assign(o);
                    return;
                }
            }

            assign(options[0]);
        }
        finally
        {
            _suppressPolicyCombo = false;
        }
    }

    [RelayCommand]
    private void CloseSidePanel()
    {
        SelectedRecord = null;
    }

    private bool CanApplySingleChange() => SelectedRecord?.IsDirty == true;

    [RelayCommand(CanExecute = nameof(CanApplySingleChange))]
    private async Task ApplySingleChangeAsync()
    {
        if (SelectedRecord is null) return;

        RunOnUi(() =>
        {
            IsPhoneBusy = true;
            PhoneStatusMessage = "Applying changes...";
            SidePanelStatusVisible = false;
            SidePanelStatusText = "";
        });

        try
        {
            var graph = RequireGraphPhone();
            var record = SelectedRecord;
            string? effectiveUserId = record.AssignmentTargetId;

            if (record.PendingUserUpn is not null &&
                record.PendingUserUpn != record.AssignmentTargetId)
            {
                RunOnUi(() => { PhoneStatusMessage = "Assigning number to user..."; });

                if (string.IsNullOrWhiteSpace(record.PendingUserUpn))
                {
                    await graph.UnassignNumberAsync(record.TelephoneNumber).ConfigureAwait(false);
                    _host.Log($"Phone Management: Unassigned {record.TelephoneNumber}");
                    effectiveUserId = null;
                }
                else
                {
                    await graph.AssignNumberAsync(record.TelephoneNumber, record.PendingUserUpn)
                        .ConfigureAwait(false);
                    _host.Log($"Phone Management: Assigned {record.TelephoneNumber} " +
                              $"to user {record.PendingUserUpn}");
                    effectiveUserId = record.PendingUserUpn;
                }
            }

            if (record.CanAssignPolicies &&
                !string.IsNullOrWhiteSpace(effectiveUserId) &&
                (!string.IsNullOrWhiteSpace(record.PendingDialPlan) ||
                 !string.IsNullOrWhiteSpace(record.PendingVoiceRoutingPolicy)))
            {
                var dialPlanGraphId = ResolveGraphPolicyId("TenantDialPlan", record.PendingDialPlan);
                var vrpGraphId = ResolveGraphPolicyId("OnlineVoiceRoutingPolicy", record.PendingVoiceRoutingPolicy);

                var dialRequested = !string.IsNullOrWhiteSpace(record.PendingDialPlan);
                var vrpRequested = !string.IsNullOrWhiteSpace(record.PendingVoiceRoutingPolicy);

                if ((dialRequested && dialPlanGraphId is null) ||
                    (vrpRequested && vrpGraphId is null))
                {
                    _host.Log("Phone Management: Could not apply — one or more Graph policy IDs " +
                              "are not cached. Click 'Load Numbers' to rebuild the cache.");
                    return;
                }

                RunOnUi(() => { PhoneStatusMessage = "Assigning policies..."; });

                await graph.AssignPoliciesAsync(effectiveUserId, dialPlanGraphId, vrpGraphId)
                    .ConfigureAwait(false);

                _host.Log($"Phone Management: Policies updated for {record.AssignedUserUpn}");
            }

            RunOnUi(() =>
            {
                record.PendingUserUpn = null;
                record.PendingDialPlan = null;
                record.PendingVoiceRoutingPolicy = null;
                record.PoliciesFetched = false;

                SidePanelStatusText = "✓ Changes applied successfully.";
                SidePanelStatusBrush = new SolidColorBrush(Colors.Green);
                SidePanelStatusVisible = true;
            });

            _host.Log($"Phone Management: Changes applied to {record.TelephoneNumber}");

            await LoadSidePanelForRecordAsync(record).ConfigureAwait(false);

            RunOnUi(() => ApplyBatchChangesCommand.NotifyCanExecuteChanged());
        }
        catch (Exception ex)
        {
            _host.Log($"Phone Management: Apply failed — {ex.Message}");
            RunOnUi(() =>
            {
                SidePanelStatusText = $"✗ Error: {ex.Message}";
                SidePanelStatusBrush = new SolidColorBrush(Colors.Red);
                SidePanelStatusVisible = true;
            });
        }
        finally
        {
            RunOnUi(() =>
            {
                IsPhoneBusy = false;
                PhoneStatusMessage = "";
                ApplySingleChangeCommand.NotifyCanExecuteChanged();
            });
        }
    }

    private bool CanApplyBatchChanges() => PhoneRecords.Any(r => r.IsDirty);

    [RelayCommand(CanExecute = nameof(CanApplyBatchChanges))]
    private async Task ApplyBatchChangesAsync()
    {
        var dirtyRecords = PhoneRecords.Where(r => r.IsDirty).ToList();
        if (dirtyRecords.Count == 0)
        {
            _host.Log("Phone Management: No pending changes to apply.");
            return;
        }

        RunOnUi(() =>
        {
            IsPhoneBusy = true;
            PhoneStatusMessage = $"Applying changes to {dirtyRecords.Count} number(s)...";
        });

        try
        {
            var graph = RequireGraphPhone();

            foreach (var record in dirtyRecords)
            {
                var effectiveUserId = record.AssignmentTargetId;

                if (record.PendingUserUpn is not null &&
                    record.PendingUserUpn != record.AssignmentTargetId)
                {
                    if (string.IsNullOrWhiteSpace(record.PendingUserUpn))
                    {
                        await graph.UnassignNumberAsync(record.TelephoneNumber).ConfigureAwait(false);
                        effectiveUserId = null;
                    }
                    else
                    {
                        await graph.AssignNumberAsync(record.TelephoneNumber, record.PendingUserUpn)
                            .ConfigureAwait(false);
                        effectiveUserId = record.PendingUserUpn;
                    }
                }

                if (record.CanAssignPolicies &&
                    !string.IsNullOrWhiteSpace(effectiveUserId) &&
                    (!string.IsNullOrWhiteSpace(record.PendingDialPlan) ||
                     !string.IsNullOrWhiteSpace(record.PendingVoiceRoutingPolicy)))
                {
                    var dialPlanGraphId = ResolveGraphPolicyId("TenantDialPlan", record.PendingDialPlan);
                    var vrpGraphId = ResolveGraphPolicyId("OnlineVoiceRoutingPolicy", record.PendingVoiceRoutingPolicy);

                    var dialRequested = !string.IsNullOrWhiteSpace(record.PendingDialPlan);
                    var vrpRequested = !string.IsNullOrWhiteSpace(record.PendingVoiceRoutingPolicy);

                    if ((dialRequested && dialPlanGraphId is null) ||
                        (vrpRequested && vrpGraphId is null))
                    {
                        _host.Log($"Phone Management: Skipped policy assignment for " +
                                  $"{record.TelephoneNumber} — Graph policy ID(s) not cached.");
                    }
                    else
                    {
                        await graph.AssignPoliciesAsync(effectiveUserId, dialPlanGraphId, vrpGraphId)
                            .ConfigureAwait(false);
                    }
                }

                record.PendingUserUpn = null;
                record.PendingDialPlan = null;
                record.PendingVoiceRoutingPolicy = null;
                record.PoliciesFetched = false;
            }

            _host.Log($"Phone Management: Applied changes to {dirtyRecords.Count} number(s).");
        }
        catch (Exception ex)
        {
            _host.Log($"Phone Management: Batch apply failed — {ex.Message}");
        }
        finally
        {
            RunOnUi(() =>
            {
                IsPhoneBusy = false;
                PhoneStatusMessage = "";
                ApplyBatchChangesCommand.NotifyCanExecuteChanged();
                GridRefreshRequested?.Invoke(this, EventArgs.Empty);
            });
        }
    }
}
