using System.Management.Automation;
using System.Management.Automation.Runspaces;

using TeamsVoiceWizard.Models;

namespace TeamsVoiceWizard.Services;

public sealed class PowerShellHost : IDisposable
{
    private readonly Runspace _runspace;
    private readonly Action<string> _log;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private readonly bool _diagnostics;
    private bool _disposed;

    public PowerShellHost(string modulePath, Action<string> log, bool diagnostics = false)
    {
        _log = log;
        _diagnostics = diagnostics;

        var iss = InitialSessionState.CreateDefault();
        _runspace = RunspaceFactory.CreateRunspace(iss);
        _runspace.Open();

        // Basic diagnostics (only when enabled)
        if (_diagnostics)
        {
            LogOutput(RunRaw(@"
                '[Startup] PSVersion: ' + $PSVersionTable.PSVersion
                '[Startup] PSHome: ' + $PSHOME
                '[Startup] HOME: ' + $HOME
                '[Startup] PSModulePath (initial): ' + $env:PSModulePath
            ", throwOnError: false));
        }

        EnsureModulePath();
        PrependLocalModulesPath(modulePath);

        if (_diagnostics)
        {
            LogOutput(RunRaw(@"'[Startup] PSModulePath (final): ' + $env:PSModulePath", throwOnError: false));
        }

        // Import dependencies explicitly
        ImportModuleChecked("Microsoft.Graph.Authentication");
        ImportModuleChecked("Microsoft.Graph.Identity.DirectoryManagement");
        ImportModuleChecked("Microsoft.Graph.Users");
        ImportModuleChecked("MicrosoftTeams");

        // Import your core module
        ImportModuleChecked(modulePath, isPath: true);

        // Verify expected functions exist
        RunRaw(@"
            Get-Command New-TVWState -ErrorAction Stop | Out-Null
            Get-Command Test-GraphConnection -ErrorAction Stop | Out-Null
            Get-Command Test-TeamsConnection -ErrorAction Stop | Out-Null
        ");

        // Expose C# log delegate to the runspace
        _runspace.SessionStateProxy.SetVariable("_csLog", log);

        // Create $state
        RunRaw(@"
            $state = New-TVWState -Country 'GB' -OnLog {
                param([string]$line)
                $_csLog.Invoke($line)
            }
        ");

        // Embed device-code helper
        RunRaw(DeviceCodeLoginScript);
    }

    private void EnsureModulePath()
    {
        var myDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);

        var ps7User = Path.Combine(myDocs, "PowerShell", "Modules");
        var ps51User = Path.Combine(myDocs, "WindowsPowerShell", "Modules");

        var ps7All = Path.Combine(programFiles, "PowerShell", "Modules");
        var ps51All = Path.Combine(programFiles, "WindowsPowerShell", "Modules");

        var ps51System = Path.Combine(windows, "System32", "WindowsPowerShell", "v1.0", "Modules");

        var script = $@"
            $paths = @(
                '{EscapePs(ps7User)}',
                '{EscapePs(ps51User)}',
                '{EscapePs(ps7All)}',
                '{EscapePs(ps51All)}',
                '{EscapePs(ps51System)}',
                ($PSHOME + '\Modules')
            ) | Where-Object {{ $_ -and (Test-Path $_) }}

            foreach ($p in $paths) {{
                if ($env:PSModulePath -notlike ('*' + $p + '*')) {{
                    $env:PSModulePath = $p + [IO.Path]::PathSeparator + $env:PSModulePath
                }}
            }}
        ";

        RunRaw(script);
    }

    private void PrependLocalModulesPath(string modulePath)
    {
        var coreDir = Path.GetDirectoryName(modulePath) ?? "";
        var localModules = Path.Combine(coreDir, "Modules");

        var script = $@"
            $p = '{EscapePs(localModules)}'
            if (Test-Path $p) {{
                if ($env:PSModulePath -notlike ('*' + $p + '*')) {{
                    $env:PSModulePath = $p + [IO.Path]::PathSeparator + $env:PSModulePath
                }}
                '[Startup] Added local module path: ' + $p
            }}
        ";

        if (_diagnostics)
        {
            LogOutput(RunRaw(script, throwOnError: false));
        }
        else
        {
            RunRaw(script, throwOnError: false);
        }
    }

    private void ImportModuleChecked(string module, bool isPath = false)
    {
        try
        {
            if (isPath)
            {
                RunRaw($"Import-Module '{EscapePs(module)}' -Force -ErrorAction Stop");
            }
            else
            {
                RunRaw($"Import-Module {module} -Force -ErrorAction Stop");
            }
        }
        catch (Exception ex)
        {
            var details = string.Join(Environment.NewLine,
                RunRaw(@"
                    $e = $Error[0]
                    '[Import-Module Diagnostics]'
                    'Message: ' + $e.Exception.Message
                    'Type: ' + $e.Exception.GetType().FullName
                    'FullyQualifiedErrorId: ' + $e.FullyQualifiedErrorId
                    'CategoryInfo: ' + $e.CategoryInfo
                    'TargetObject: ' + ($e.TargetObject | Out-String)
                    'InvocationInfo: ' + ($e.InvocationInfo | Out-String)
                    'ScriptStackTrace: ' + ($e.ScriptStackTrace | Out-String)
                ", throwOnError: false).Select(o => o?.ToString() ?? ""));

            _log($"[PS Error] Import-Module failed: {module}");
            if (_diagnostics) _log(details);

            if (!isPath && _diagnostics)
            {
                var found = string.Join(Environment.NewLine,
                    RunRaw($@"
                        $m = Get-Module -ListAvailable {module} | Sort-Object Version -Descending | Select-Object -First 3
                        if (-not $m) {{ 'Get-Module -ListAvailable found nothing for {module}' }}
                        else {{ $m | Select Name, Version, Path | Format-List | Out-String }}
                    ", throwOnError: false).Select(o => o?.ToString() ?? ""));

                _log(found);
            }

            throw new InvalidOperationException($"Module load failed: {module}", ex);
        }
    }

    // Public API
    public Task ConnectGraphAsync(IEnumerable<string> scopes) =>
        RunAsync($@"
            $scopeList = @({string.Join(",", scopes.Select(s => $"'{s}'"))})
            Invoke-TVWDeviceCodeLogin -State $state -Scopes $scopeList
        ");

    public Task ConnectTeamsAsync() =>
        RunAsync("Connect-MicrosoftTeams | Out-Null");

    public Task<bool> IsGraphConnectedAsync() =>
        RunScalarAsync<bool>("Test-GraphConnection");

    public Task<bool> IsTeamsConnectedAsync() =>
        RunScalarAsync<bool>("Test-TeamsConnection");

    public Task<(List<string> Domains, bool DerivedDetected, string OutputText)>
        CreateDomainsAndGetTxtAsync(string endpointInput, string site, string country) =>
        Task.Run(async () =>
        {
            await _lock.WaitAsync().ConfigureAwait(false);
            try
            {
                SetVar("_endpointInput", endpointInput);
                SetVar("_site", site);
                SetVar("_country", country);

                RunRaw(@"
                    $state.Site    = $_site
                    $state.Country = $_country

                    $derivedRef = [ref]$false
                    $domains = Get-PstnGatewaysFromInput `
                        -EndpointInput $_endpointInput `
                        -DerivedTrunkDetected $derivedRef

                    $state.Domains           = $domains
                    $state.DerivedTrunkModel = $derivedRef.Value

                    $domainResults = New-GammaDomains -Domains $domains
                    $txtRecords    = Get-GammaDomainTxtRecords -Domains $domains
                    $state.TxtRecords = $txtRecords
                ");

                var domains = ToStringList(RunRaw("$state.Domains"));

                var derivedRaw = RunRaw("$state.DerivedTrunkModel").FirstOrDefault()?.BaseObject;
                var derived = derivedRaw is bool b
                    ? b
                    : (derivedRaw is not null && bool.TryParse(derivedRaw.ToString(), out var parsed) ? parsed : false);

                var outText = string.Join(Environment.NewLine,
                    RunRaw(@"
                        'Domains:'
                        $domainResults | Format-Table -AutoSize | Out-String
                        ''
                        'TXT Records:'
                        $txtRecords | Format-Table -AutoSize | Out-String
                    ").Select(o => o?.ToString() ?? ""));

                return (domains, derived, outText);
            }
            finally { _lock.Release(); }
        });

    public Task<List<VerificationResult>> VerifyDomainsAsync() =>
        Task.Run(async () =>
        {
            await _lock.WaitAsync().ConfigureAwait(false);
            try
            {
                var raw = RunRaw(@"
                    $v = Test-GammaDomainsVerification -Domains $state.Domains
                    $state.Verification = $v
                    $v
                ");

                return raw.Select(o =>
                {
                    var domain = o.Properties["Domain"]?.Value?.ToString() ?? "";
                    var verified = o.Properties["Verified"]?.Value is true;
                    var error = o.Properties["Error"]?.Value?.ToString();
                    return new VerificationResult(domain, verified, error);
                }).ToList();
            }
            finally { _lock.Release(); }
        });

    public Task<List<SkuRecord>> LoadLicenseInventoryAsync(int minFree = 2) =>
        Task.Run(async () =>
        {
            await _lock.WaitAsync().ConfigureAwait(false);
            try
            {
                SetVar("_minFree", minFree);

                var raw = RunRaw(@"
                    Get-LicenseInventory -MinimumFree $_minFree |
                        Where-Object { $_.ExistsInTenant -and $_.MeetsMinFree }
                ");

                return raw.Select(o =>
                {
                    Guid? skuId = null;
                    var rawSku = o.Properties["SkuId"]?.Value?.ToString();
                    if (Guid.TryParse(rawSku, out var g)) skuId = g;

                    return new SkuRecord(
                        Family: o.Properties["Family"]?.Value?.ToString() ?? "",
                        Product: o.Properties["Product"]?.Value?.ToString() ?? "",
                        SkuPartNumber: o.Properties["SkuPartNumber"]?.Value?.ToString() ?? "",
                        SkuId: skuId,
                        TotalEnabled: Convert.ToInt32(o.Properties["TotalEnabled"]?.Value ?? 0),
                        Consumed: Convert.ToInt32(o.Properties["Consumed"]?.Value ?? 0),
                        Available: Convert.ToInt32(o.Properties["Available"]?.Value ?? 0),
                        MeetsMinFree: o.Properties["MeetsMinFree"]?.Value is true,
                        ExistsInTenant: o.Properties["ExistsInTenant"]?.Value is true
                    );
                }).ToList();
            }
            finally { _lock.Release(); }
        });

    public Task<(List<string> Users, List<string> ResourceAccounts)>
        CreateTestObjectsAsync(SkuRecord sku, string country) =>
        Task.Run(async () =>
        {
            await _lock.WaitAsync().ConfigureAwait(false);
            try
            {
                SetVar("_skuId", sku.SkuId?.ToString() ?? "");
                SetVar("_skuPartNumber", sku.SkuPartNumber);
                SetVar("_country", country);

                RunRaw(@"
                    $state.ChosenSku = $_skuPartNumber

                    $raSkuPart = 'PHONESYSTEM_VIRTUALUSER'

                    foreach ($gw in $state.Domains) {
                        if ($_skuPartNumber -eq $raSkuPart) {
                            $raUpn = Create-ResourceAccounts -PSTNGW $gw
                            $state.CreatedResourceAccounts += $raUpn
                        } else {
                            $upns = New-StandardTestUsers -PSTNGW $gw -UsageLocation $_country
                            $state.CreatedUsers += $upns
                            $skuGuid = [guid]$_skuId
                            foreach ($upn in $upns) {
                                Assign-LicenseSku -UserPrincipalName $upn -SkuId $skuGuid | Out-Null
                            }
                        }
                    }
                ");

                var users = ToStringList(RunRaw("$state.CreatedUsers"));
                var ras = ToStringList(RunRaw("$state.CreatedResourceAccounts"));

                return (users, ras);
            }
            finally { _lock.Release(); }
        });

    public Task CleanupTestObjectsAsync(List<string> userUpns, List<string> raUpns) =>
        Task.Run(async () =>
        {
            await _lock.WaitAsync().ConfigureAwait(false);
            try
            {
                SetVar("_userUpns", userUpns.ToArray());
                SetVar("_raUpns", raUpns.ToArray());

                RunRaw(@"
                    $results = Remove-CreatedTestObjects -UserUpns $_userUpns -ResourceUpns $_raUpns
                    foreach ($r in $results) {
                        Write-TVWLog -State $state -Text $r
                    }
                    $state.CreatedUsers             = @()
                    $state.CreatedResourceAccounts  = @()
                ");
            }
            finally { _lock.Release(); }
        });

    public Task ApplyVoiceConfigurationAsync(string site, string country, bool derivedTrunkModel) =>
        Task.Run(async () =>
        {
            await _lock.WaitAsync().ConfigureAwait(false);
            try
            {
                SetVar("_site", site);
                SetVar("_country", country);
                SetVar("_derived", derivedTrunkModel);

                RunRaw(@"
                    New-UKVoiceConfiguration `
                        -Site               $_site `
                        -Country            $_country `
                        -PstnGateways       $state.Domains `
                        -DerivedTrunkModel  $_derived
                ");
            }
            finally { _lock.Release(); }
        });

    private ICollection<PSObject> RunRaw(string script, bool throwOnError = true)
    {
        using var ps = PowerShell.Create();
        ps.Runspace = _runspace;
        ps.AddScript(script);

        var results = ps.Invoke(); // synchronous call

        if (ps.Streams.Error.Count > 0)
        {
            var msg = string.Join(Environment.NewLine, ps.Streams.Error.Select(e => e.ToString()));
            _log($"[PS Error] {msg}");

            if (throwOnError)
                throw new InvalidOperationException(msg);
        }

        return results;
    }

    private async Task RunAsync(string script)
    {
        await _lock.WaitAsync().ConfigureAwait(false);
        try { await Task.Run(() => RunRaw(script)).ConfigureAwait(false); }
        finally { _lock.Release(); }
    }

    private async Task<T> RunScalarAsync<T>(string script)
    {
        await _lock.WaitAsync().ConfigureAwait(false);
        try
        {
            var results = await Task.Run(() => RunRaw(script)).ConfigureAwait(false);
            var first = results.FirstOrDefault();

            if (first?.BaseObject is T typed) return typed;

            try { return (T)Convert.ChangeType(first?.BaseObject, typeof(T))!; }
            catch { return default!; }
        }
        finally { _lock.Release(); }
    }

    // ✅ SetVar reintroduced here (internal helpers region)
    private void SetVar(string name, object? value) =>
        _runspace.SessionStateProxy.SetVariable(name, value);

    private static string EscapePs(string path) => path.Replace("'", "''");

    private static List<string> ToStringList(IEnumerable<PSObject> raw) =>
        raw.Select(o => o?.BaseObject?.ToString() ?? o?.ToString() ?? "")
           .Select(s => s.Trim())
           .Where(s => !string.IsNullOrWhiteSpace(s))
           .ToList();

    private void LogOutput(ICollection<PSObject> output)
    {
        foreach (var o in output)
        {
            var s = o?.ToString();
            if (!string.IsNullOrWhiteSpace(s))
                _log(s!);
        }
    }

    private const string DeviceCodeLoginScript = @"
function Invoke-TVWDeviceCodeLogin {
    param(
        [Parameter(Mandatory)]$State,
        [string[]]$Scopes
    )

    $clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
    $tenant   = 'organizations'

    $scopeStr = ($Scopes | ForEach-Object { ""https://graph.microsoft.com/$_"" }) -join ' '
    $scopeStr = ""$scopeStr offline_access openid profile""

    $deviceCodeResp = Invoke-RestMethod -Method POST `
        -Uri ""https://login.microsoftonline.com/$tenant/oauth2/v2.0/devicecode"" `
        -Body @{ client_id = $clientId; scope = $scopeStr }

    Write-TVWLog -State $State -Text $deviceCodeResp.message

    $tokenUri  = ""https://login.microsoftonline.com/$tenant/oauth2/v2.0/token""
    $expiresAt = (Get-Date).AddSeconds([int]$deviceCodeResp.expires_in)
    $interval  = [int]$deviceCodeResp.interval

    while ((Get-Date) -lt $expiresAt) {
        $resp = Invoke-RestMethod -Method POST -Uri $tokenUri `
            -Body @{
                grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id   = $clientId
                device_code = $deviceCodeResp.device_code
            } -SkipHttpErrorCheck -StatusCodeVariable status

        if ($status -eq 200 -and $resp.access_token) {
            $secureToken = ConvertTo-SecureString $resp.access_token -AsPlainText -Force
            Connect-MgGraph -AccessToken $secureToken -NoWelcome
            return $true
        }

        switch ($resp.error) {
            'authorization_pending' { Start-Sleep -Seconds $interval;        continue }
            'slow_down'             { Start-Sleep -Seconds ($interval + 5);  continue }
            'access_denied'         { throw ""User cancelled sign-in."" }
            'expired_token'         { throw ""Device code expired."" }
            default                 { throw ""Token polling failed: $($resp.error) $($resp.error_description)"" }
        }
    }
    throw 'Device code expired before authentication completed.'
}
";

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _lock.Dispose();
        _runspace.Dispose();
    }
}