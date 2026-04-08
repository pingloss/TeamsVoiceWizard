using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.UI.Dispatching;
using TeamsVoiceWizard.Models;
using TeamsVoiceWizard.Services;

namespace TeamsVoiceWizard.ViewModels;

/// <summary>
/// Drives the Bulk Import ContentDialog.
/// Created fresh each time the dialog is opened; owns its own cancellation lifecycle.
/// </summary>
public partial class BulkImportViewModel : ObservableObject
{
    private readonly GraphPhoneService              _graph;
    private readonly IReadOnlyList<PhoneNumberRecord> _tenantNumbers;
    private readonly IReadOnlyList<PolicyEntry>     _dialPlans;
    private readonly IReadOnlyList<PolicyEntry>     _vrPolicies;
    private readonly Dictionary<string, string>     _graphPolicyIdCache;
    private readonly Action<string>                 _log;
    private readonly DispatcherQueue                _dispatcher;

    private CancellationTokenSource? _cts;

    public BulkImportViewModel(
        GraphPhoneService              graph,
        IReadOnlyList<PhoneNumberRecord> tenantNumbers,
        IReadOnlyList<PolicyEntry>     dialPlans,
        IReadOnlyList<PolicyEntry>     vrPolicies,
        Dictionary<string, string>     graphPolicyIdCache,
        Action<string>                 log)
    {
        _graph              = graph;
        _tenantNumbers      = tenantNumbers;
        _dialPlans          = dialPlans;
        _vrPolicies         = vrPolicies;
        _graphPolicyIdCache = graphPolicyIdCache;
        _log                = log;
        _dispatcher         = DispatcherQueue.GetForCurrentThread();
    }

    // ── Observable state ──────────────────────────────────────────────────────

    public ObservableCollection<BulkImportRow> Rows { get; } = new();

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasRows))]
    [NotifyPropertyChangedFor(nameof(StatusSummary))]
    private string _fileName = "No file selected";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasRows))]
    private bool _isLoading;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanApply))]
    private bool _isValidating;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanApply))]
    private bool _isApplying;

    [ObservableProperty]
    private string _progressMessage = "";

    [ObservableProperty]
    private string _statusSummary = "";

    [ObservableProperty]
    private bool _validationComplete;

    public bool HasRows    => Rows.Count > 0 && !IsLoading;
    public bool CanValidate => HasRows && !IsValidating && !IsApplying;
    public bool CanApply   => HasRows && !IsValidating && !IsApplying &&
                              Rows.Any(r => r.Status == BulkImportStatus.Valid);

    // Raised when apply completes so the dialog can update its primary button
    public event EventHandler? ApplyCompleted;

    // ── File loading ──────────────────────────────────────────────────────────

    /// <summary>
    /// Parses the given file stream. Called from the dialog code-behind after the
    /// file picker returns. <paramref name="extension"/> should be ".csv" or ".xlsx".
    /// </summary>
    public async Task LoadFileAsync(Stream stream, string fileName, string extension)
    {
        IsLoading    = true;
        FileName     = fileName;
        StatusSummary = "";
        ValidatedComplete();

        try
        {
            Rows.Clear();

            List<BulkImportRow> parsed = await Task.Run(() =>
                extension.Equals(".csv", StringComparison.OrdinalIgnoreCase)
                    ? BulkImportParser.ParseCsv(stream)
                    : BulkImportParser.ParseXlsx(stream)
            ).ConfigureAwait(true);

            // Structural validation is fast — run synchronously on UI thread
            BulkImportValidator.ValidateStructure(parsed);

            foreach (var row in parsed)
                Rows.Add(row);

            int structErrors = parsed.Count(r => r.Status == BulkImportStatus.Error);
            StatusSummary = structErrors > 0
                ? $"{parsed.Count} row(s) loaded — {structErrors} structural error(s). Fix before validating."
                : $"{parsed.Count} row(s) loaded. Click 'Validate' to check against tenant.";

            _log($"Bulk Import: Parsed {parsed.Count} row(s) from '{fileName}'.");
        }
        catch (Exception ex)
        {
            StatusSummary = $"Parse error: {ex.Message}";
            _log($"Bulk Import: Parse failed — {ex.Message}");
        }
        finally
        {
            IsLoading = false;
            NotifyCommandsChanged();
        }
    }

    // ── Validation ────────────────────────────────────────────────────────────

    [RelayCommand]
    private async Task ValidateAsync()
    {
        if (!CanValidate) return;

        _cts = new CancellationTokenSource();
        IsValidating      = true;
        ValidationComplete = false;
        ProgressMessage   = "Validating rows against tenant…";
        StatusSummary     = "";

        try
        {
            await BulkImportValidator.ValidateTenantAsync(
                Rows,
                _graph,
                _tenantNumbers,
                _dialPlans,
                _vrPolicies,
                _graphPolicyIdCache,
                onProgress: row => RunOnUi(() => RefreshRow(row)),
                ct: _cts.Token
            ).ConfigureAwait(true);

            int valid   = Rows.Count(r => r.Status == BulkImportStatus.Valid);
            int errors  = Rows.Count(r => r.Status == BulkImportStatus.Error);
            int pending = Rows.Count(r => r.Status == BulkImportStatus.Pending);

            StatusSummary = pending > 0
                ? $"{valid} valid, {errors} error(s), {pending} skipped (structural errors)."
                : $"{valid} valid, {errors} error(s).";

            ValidationComplete = true;
            _log($"Bulk Import: Validation complete — {valid} valid, {errors} error(s).");
        }
        catch (OperationCanceledException)
        {
            StatusSummary = "Validation cancelled.";
        }
        catch (Exception ex)
        {
            StatusSummary = $"Validation error: {ex.Message}";
            _log($"Bulk Import: Validation failed — {ex.Message}");
        }
        finally
        {
            IsValidating    = false;
            ProgressMessage = "";
            NotifyCommandsChanged();
        }
    }

    // ── Apply ─────────────────────────────────────────────────────────────────

    [RelayCommand]
    private async Task ApplyAsync()
    {
        if (!CanApply) return;

        var validRows = Rows.Where(r => r.Status == BulkImportStatus.Valid).ToList();
        if (validRows.Count == 0) return;

        _cts = new CancellationTokenSource();
        IsApplying      = true;
        ProgressMessage = $"Applying {validRows.Count} row(s)…";

        int applied = 0, failed = 0;

        try
        {
            foreach (var row in validRows)
            {
                _cts.Token.ThrowIfCancellationRequested();

                RunOnUi(() => ProgressMessage = $"Applying row {row.RowNumber} ({row.PhoneNumber})…");

                try
                {
                    // 1. Assign number
                    await _graph.AssignNumberAsync(row.PhoneNumber, row.ResolvedUserId!)
                        .ConfigureAwait(false);

                    // 2. Assign policies if Direct Routing
                    if (row.RequiresPolicies &&
                        (!string.IsNullOrWhiteSpace(row.ResolvedDialPlanGraphId) ||
                         !string.IsNullOrWhiteSpace(row.ResolvedVrpGraphId)))
                    {
                        await _graph.AssignPoliciesAsync(
                            row.ResolvedUserId!,
                            row.ResolvedDialPlanGraphId,
                            row.ResolvedVrpGraphId).ConfigureAwait(false);
                    }

                    RunOnUi(() =>
                    {
                        row.Status            = BulkImportStatus.Applied;
                        row.ValidationMessage = "";
                    });

                    applied++;
                    _log($"Bulk Import: Applied {row.PhoneNumber} → {row.UserPrincipalName}.");
                }
                catch (Exception ex)
                {
                    RunOnUi(() =>
                    {
                        row.Status            = BulkImportStatus.Error;
                        row.ValidationMessage = ex.Message;
                    });

                    failed++;
                    _log($"Bulk Import: Row {row.RowNumber} failed — {ex.Message}");
                }
            }

            StatusSummary = $"Apply complete: {applied} applied, {failed} failed.";
            _log($"Bulk Import: Apply complete — {applied} applied, {failed} failed.");
        }
        catch (OperationCanceledException)
        {
            StatusSummary = $"Apply cancelled. {applied} applied before cancellation.";
        }
        finally
        {
            IsApplying      = false;
            ProgressMessage = "";
            NotifyCommandsChanged();
            ApplyCompleted?.Invoke(this, EventArgs.Empty);
        }
    }

    // ── Cancel ────────────────────────────────────────────────────────────────

    [RelayCommand]
    private void Cancel()
    {
        _cts?.Cancel();
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private void ValidatedComplete() => ValidationComplete = false;

    private void RefreshRow(BulkImportRow row)
    {
        // Force the DataGrid to notice the status change by triggering a collection notification.
        // The row implements INotifyPropertyChanged so this is automatic via binding.
        // This no-op is intentional — INotifyPropertyChanged on the row handles UI refresh.
        _ = row;
    }

    private void RunOnUi(Action action)
    {
        if (_dispatcher.HasThreadAccess)
            action();
        else
            _dispatcher.TryEnqueue(() => action());
    }

    private void NotifyCommandsChanged()
    {
        OnPropertyChanged(nameof(HasRows));
        OnPropertyChanged(nameof(CanValidate));
        OnPropertyChanged(nameof(CanApply));
        ValidateCommand.NotifyCanExecuteChanged();
        ApplyCommand.NotifyCanExecuteChanged();
    }
}
