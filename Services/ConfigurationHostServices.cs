using TeamsVoiceWizard.Models;

namespace TeamsVoiceWizard.Services;

/// <summary>
/// Bridges MainWindow (PowerShell host, Graph token, logging) into the Configuration tab MVVM.
/// All delegates close over MainWindow's private fields so the ViewModel never references the window directly.
/// </summary>
public sealed record ConfigurationHostServices(
    Func<bool>                                                                              IsPowerShellReady,
    Func<string[], Task>                                                                    ConnectGraphAsync,
    Func<Task>                                                                              ConnectTeamsAsync,
    Func<Task<bool>>                                                                        IsGraphConnectedAsync,
    Func<Task<bool>>                                                                        IsTeamsConnectedAsync,
    Func<string, string, string, Task<(List<string> Domains, bool DerivedDetected, string OutputText)>>
                                                                                            CreateDomainsAndGetTxtAsync,
    Func<Task<List<VerificationResult>>>                                                    VerifyDomainsAsync,
    Func<int, Task<List<SkuRecord>>>                                                        LoadLicenseInventoryAsync,
    Func<SkuRecord, string, Task<(List<string> Users, List<string> ResourceAccounts)>>     CreateTestObjectsAsync,
    Func<List<string>, List<string>, Task>                                                  CleanupTestObjectsAsync,
    Func<string, string, bool, Task>                                                        ApplyVoiceConfigurationAsync
);
