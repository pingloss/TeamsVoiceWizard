# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

TeamsVoiceWizard is a Windows desktop application (WinUI 3) for managing Microsoft Teams Phone System voice configurations and PSTN gateway provisioning. It targets enterprise Teams administrators.

## Build Commands

```bash
dotnet restore                  # Restore NuGet packages
dotnet build                    # Debug build
dotnet build -c Release         # Release build
dotnet build --arch x64         # x64-specific build
```

Build output goes to `~/VSBuilds/TeamsVoiceWizard/bin/` (configured in `Directory.Build.props`).

There are no automated tests in this project.

## Tech Stack

- **UI:** WinUI 3 (Windows App SDK 1.8)
- **Language:** C# 12, .NET 9.0-windows10.0.19041
- **MVVM:** CommunityToolkit.Mvvm 8.4.0
- **Data Grid:** CommunityToolkit.WinUI.UI.Controls.DataGrid
- **Graph API:** Microsoft Graph SDK v2.32.0+
- **PowerShell integration:** System.Management.Automation (embedded runspace)
- **Document parsing:** DocumentFormat.OpenXml (CSV/XLSX bulk import)

## Architecture

**Pattern:** MVVM with layered service architecture.

### Layers

**Views** (`Views/`) — WinUI 3 XAML pages:
- `ConfigurationView` — PSTN gateway setup, domain creation, policy application
- `PhoneManagementView` — phone number assignment, policy management, DataGrid
- `BulkImportDialog` — CSV/XLSX import wizard

**ViewModels** (`ViewModels/`) — CommunityToolkit.Mvvm `ObservableObject` subclasses:
- State, commands, and orchestration logic
- Commands use `[RelayCommand]` attributes
- Property notifications use `[ObservableProperty]` and `[NotifyPropertyChangedFor]`

**Services** (`Services/`):
- `PowerShellHost.cs` — C# wrapper around a PowerShell 7 runspace; runs the embedded `core/TeamsVoiceWizard.Core.psm1` module, manages device-code authentication, module imports, and infrastructure commands
- `GraphPhoneService.cs` — Microsoft Graph API client for phone number CRUD, user resolution, policy assignment, and paginated queries
- `BulkImportParser.cs` / `BulkImportValidator.cs` — CSV/XLSX parsing and validation

**Models** (`Models/`):
- `WizardState.cs` — session-level state shared across the app (domains, users, licenses, connections)
- `PhoneNumberRecord.cs` — individual number with assignment metadata
- `PolicyCaches.cs` — in-memory cache for dial plans and voice routing policies

**PowerShell Module** (`core/`):
- `TeamsVoiceWizard.Core.psm1` — ~500-line module handling domain creation, verification, license inventory, test object lifecycle, and voice config application
- Copied to output directory at build time (see `.csproj`)

**MainWindow** (`MainWindow.xaml.cs`) — app shell:
- Tab navigation between Configuration and Phone Management
- Wires ViewModels to service bridges via sealed records (`ConfigurationHostServices`, `PhoneManagementHostServices`) containing function delegates — this avoids direct window references from ViewModels
- Runs DPI-aware window sizing on startup
- Coordinates `InitPowerShellAsync()` startup sequence

### Key Patterns

**Service bridging:** ViewModels receive a sealed record of delegate functions (not direct service references). Defined in `Services/` as `ConfigurationHostServices` and `PhoneManagementHostServices`.

**UI thread marshaling:** All UI updates from async operations go through `DispatcherQueue.TryEnqueue(...)`.

**Buffered logging:** Log output accumulates in a `StringBuilder`, then is flushed to the UI in a single assignment to avoid O(n²) string concatenation.

**Policy caching:** `PolicyCaches` holds in-memory dial plan and voice routing policy lists, populated once per session.

**Concurrency control:** `SemaphoreSlim` locks protect operations that must not run concurrently (e.g., bulk Graph updates).

**PowerShell state:** A `$global:_tvwState` hashtable on the PowerShell side mirrors the C# `WizardState`. Graph token stored as `$global:_tvwGraphToken`.

### Startup Sequence

1. `App.OnLaunched()` → creates `MainWindow`
2. `MainWindow_Loaded` → `WireViewModels()`, then `InitPowerShellAsync()`
3. `InitPowerShellAsync()` loads the `.psm1` module, imports Teams/Graph modules, creates PS state object, tests Graph and Teams connections
4. Phone Management tab is enabled only after Graph connection succeeds

## WinUI / DataGrid Notes

- `SortMemberPath` is not supported in the Uno DataGrid port (if migrating). Use `Tag`-based workaround instead.
- `DataGrid` column sorting in `PhoneManagementView` was added in the most recent commits — see `PhoneManagementView.xaml` and `PhoneManagementViewModel.cs` for the pattern.
