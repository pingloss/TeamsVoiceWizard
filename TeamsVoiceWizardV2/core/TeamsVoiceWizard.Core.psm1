Set-StrictMode -Version Latest

# ============================================================
# State & Logging
# ============================================================

function New-TVWState {
    param(
        [string]$Country = "GB",
        [scriptblock]$OnLog
    )

    [pscustomobject]@{
        GammaEndpoint           = $null
        Site                    = $null
        Country                 = $Country
        Domains                 = @()
        TxtRecords              = @()
        Verification            = @()
        LicenseInventory        = @()
        ChosenSku               = $null
        CreatedUsers            = @()
        CreatedResourceAccounts = @()

        DerivedTrunkModel       = $true

        Log                     = New-Object System.Collections.Generic.List[string]
        OnLog                   = $OnLog
    }
}

function Write-TVWLog {
    param(
        [Parameter(Mandatory)]$State,
        [Parameter(Mandatory)][string]$Text
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "$timestamp $Text"
    $State.Log.Add($line) | Out-Null

    if ($State.OnLog) {
        & $State.OnLog $line
    }
    else {
        Write-Host $line
    }
}

# ============================================================
# Module Management (STRICT)
# ============================================================
# - This module refuses to load unless required dependencies are importable at/above minimum versions.
# - Module discovery depends on $env:PSModulePath which varies by host/start method. [1](https://stackoverflow.com/questions/78009842/uno-platform-setup)
# - Some Graph module behaviours can differ in embedded hosts; the error details help diagnose. [4](https://github.com/unoplatform/uno/issues/18281)

$script:TVW_RequiredModules = @(
    @{ Name = "Microsoft.Graph.Identity.DirectoryManagement"; MinimumVersion = [Version]"2.32.0" },
    @{ Name = "Microsoft.Graph.Authentication";               MinimumVersion = [Version]"2.32.0" },
    @{ Name = "Microsoft.Graph.Users";                        MinimumVersion = [Version]"2.32.0" },
    @{ Name = "MicrosoftTeams";                               MinimumVersion = [Version]"6.7.0"  }
)

function Get-TVWModuleBestCandidate {
    param([Parameter(Mandatory)][string]$Name)

    Get-Module -ListAvailable -Name $Name |
        Sort-Object Version -Descending |
        Select-Object -First 1
}

function Assert-TVWDependencies {
    [CmdletBinding()]
    param()

    $errors = New-Object System.Collections.Generic.List[string]

    foreach ($m in $script:TVW_RequiredModules) {
        $name = $m.Name
        $min  = $m.MinimumVersion

        $candidate = Get-TVWModuleBestCandidate -Name $name

        if (-not $candidate) {
            $errors.Add("Missing module '$name' (minimum required: $min).")
            continue
        }

        if ([Version]$candidate.Version -lt $min) {
            $errors.Add("Module '$name' found ($($candidate.Version)) but minimum required is $min. Path: $($candidate.Path)")
            continue
        }

        try {
            Import-Module $name -MinimumVersion $min -Force -ErrorAction Stop
        }
        catch {
            $detail = $_ | Out-String
            $err0   = $Error[0] | Format-List * -Force | Out-String

            $errors.Add(@"
Module '$name' was found (version $($candidate.Version)) but failed to import.
Path: $($candidate.Path)
Exception: $($_.Exception.GetType().FullName) - $($_.Exception.Message)

--- Error Record ---
$detail

--- Error[0] (expanded) ---
$err0
"@.Trim())
        }
    }

    if ($errors.Count -gt 0) {
        $msg = @()
        $msg += "TeamsVoiceWizard.Core.psm1 dependency check failed. STRICT mode is enabled and the module will not load."
        $msg += ""
        $msg += "Failures:"
        $msg += ($errors | ForEach-Object { " - " + $_ })
        $msg += ""
        $msg += "PSModulePath (searched locations):"
        $msg += $env:PSModulePath
        $msg += ""
        $msg += "Fix guidance:"
        $msg += " - Ensure required modules are installed and importable in this PowerShell host context."
        $msg += " - If this fails only when hosted in-app (but works in pwsh), inspect the exception details above for assembly/type load conflicts."

        throw ($msg -join [Environment]::NewLine)
    }
}

Assert-TVWDependencies

# ============================================================
# Connection Helpers
# ============================================================

function Test-GraphConnection {
    try {
        Get-MgDomain -ErrorAction Stop | Select-Object -First 1 | Out-Null
        return $true
    }
    catch { return $false }
}

function Test-TeamsConnection {
    try {
        Get-CsTenantDialPlan -ErrorAction Stop | Out-Null
        return $true
    }
    catch { return $false }
}

# ============================================================
# Gamma / Domain Management
# ============================================================

function Get-PstnGatewaysFromGammaEndpoint {
    param([Parameter(Mandatory)][string]$GammaEndpoint)

    @("$GammaEndpoint.customers.voiceconnected.co.uk")
}

function Get-PstnGatewaysFromInput {
    param(
        [Parameter(Mandatory)][string]$EndpointInput,
        [ref]$DerivedTrunkDetected
    )

    $DerivedTrunkDetected.Value = $false
    $gateways = @()

    $values = $EndpointInput -split ',' |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ }

    foreach ($val in $values) {

        if ($val -match '\.') {
            $gateways += $val
            if ($val -match '\.ucconnect\.co\.uk$') {
                $DerivedTrunkDetected.Value = $true
            }
            continue
        }

        $gateways += "$val.part03.ucconnect.co.uk"
        $gateways += "$val.part13.ucconnect.co.uk"
        $DerivedTrunkDetected.Value = $true
    }

    $gateways | Sort-Object -Unique
}

function New-GammaDomains {
    param([string[]]$Domains)

    foreach ($domainName in $Domains) {
        try {
            New-MgDomain -Id $domainName -IsDefault:$false -WarningAction SilentlyContinue | Out-Null
            [pscustomobject]@{ Domain = $domainName; Status = "CreatedOrExists" }
        }
        catch {
            [pscustomobject]@{ Domain = $domainName; Status = "Error: $($_.Exception.Message)" }
        }
    }
}

function Get-GammaDomainTxtRecords {
    param([string[]]$Domains)

    foreach ($domainName in $Domains) {
        try {
            $txt = (Get-MgDomainVerificationDnsRecord -DomainId $domainName |
                    Where-Object RecordType -eq "Txt").AdditionalProperties.text

            [pscustomobject]@{
                Domain   = $domainName
                Txt      = $txt
                Verified = $false
                Error    = $null
            }
        }
        catch {
            [pscustomobject]@{
                Domain   = $domainName
                Txt      = $null
                Verified = $false
                Error    = $_.Exception.Message
            }
        }
    }
}

function Test-GammaDomainsVerification {
    param([string[]]$Domains)

    foreach ($domainName in $Domains) {
        try {
            Confirm-MgDomain -DomainId $domainName -WarningAction SilentlyContinue | Out-Null
            [pscustomobject]@{ Domain = $domainName; Verified = $true; Error = $null }
        }
        catch {
            [pscustomobject]@{ Domain = $domainName; Verified = $false; Error = $_.Exception.Message }
        }
    }
}

# ============================================================
# Licensing / SKUs
# ============================================================

function Normalize-SkuPartNumber {
    param([Parameter(Mandatory)][string]$Value)
    ($Value -replace "[\u200B\u200C\u200D\uFEFF]", "").Trim()
}

function Get-CommercialTargetSkus {
    @(
        [pscustomobject]@{ Family="Enterprise"; Product="Office 365 E1"; SkuPartNumber="STANDARDPACK" },
        [pscustomobject]@{ Family="Enterprise"; Product="Office 365 E3"; SkuPartNumber="ENTERPRISEPACK" },
        [pscustomobject]@{ Family="Enterprise"; Product="Office 365 E5"; SkuPartNumber="ENTERPRISEPREMIUM" },
        [pscustomobject]@{ Family="Enterprise"; Product="Microsoft 365 E3"; SkuPartNumber="SPE_E3" },
        [pscustomobject]@{ Family="Enterprise"; Product="Microsoft 365 E5"; SkuPartNumber="SPE_E5" },
        [pscustomobject]@{ Family="Frontline";  Product="Office 365 F3";  SkuPartNumber="DESKLESSPACK" },
        [pscustomobject]@{ Family="Frontline";  Product="Microsoft 365 F1";SkuPartNumber="M365_F1" },
        [pscustomobject]@{ Family="Frontline";  Product="Microsoft 365 F3";SkuPartNumber="SPE_F3" },
        [pscustomobject]@{ Family="Business";   Product="M365 Business Basic";   SkuPartNumber="O365_BUSINESS_ESSENTIALS" },
        [pscustomobject]@{ Family="Business";   Product="M365 Business Standard";SkuPartNumber="O365_BUSINESS_PREMIUM" },
        [pscustomobject]@{ Family="Business";   Product="M365 Business Premium"; SkuPartNumber="SPB" },
        [pscustomobject]@{ Family="Teams Add-ons"; Product="Teams Shared Devices";SkuPartNumber="MCOCAP" },
        [pscustomobject]@{ Family="Teams Add-ons"; Product="Teams Phone Resource Account";SkuPartNumber="PHONESYSTEM_VIRTUALUSER" }
    )
}

function Get-LicenseInventory {
    param([int]$MinimumFree = 2)

    $skus = Get-MgSubscribedSku -All
    $index = @{}

    foreach ($sku in $skus) {
        $index[(Normalize-SkuPartNumber $sku.SkuPartNumber)] = $sku
    }

    foreach ($t in Get-CommercialTargetSkus) {
        $match = $index[$t.SkuPartNumber]

        $enabled  = if ($match) { [int]$match.PrepaidUnits.Enabled } else { 0 }
        $consumed = if ($match) { [int]$match.ConsumedUnits } else { 0 }
        $free     = $enabled - $consumed

        [pscustomobject]@{
            Family         = $t.Family
            Product        = $t.Product
            SkuPartNumber  = $t.SkuPartNumber
            SkuId          = if ($match) { $match.SkuId } else { $null }
            TotalEnabled   = $enabled
            Consumed       = $consumed
            Available      = $free
            MeetsMinFree   = ($free -ge $MinimumFree)
            ExistsInTenant = [bool]$match
        }
    }
}

function Assign-LicenseSku {
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [Parameter(Mandatory)][guid]$SkuId
    )

    try {
        Set-MgUserLicense -UserId $UserPrincipalName `
            -AddLicenses @(@{SkuId = $SkuId}) `
            -RemoveLicenses @() `
            -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        Write-Error "License assign failed for ${UserPrincipalName}: $($_.Exception.Message)"
        return $false
    }
}

# ============================================================
# Test Objects
# ============================================================

function Create-ResourceAccounts {
    param([Parameter(Mandatory)][string]$PSTNGW)

    $upn = "testuser@$PSTNGW"

    try {
        New-CsOnlineApplicationInstance `
            -UserPrincipalName $upn `
            -ApplicationId "ce933385-9390-45d1-9512-c8d228074e07" `
            -DisplayName "Test Resource Account" `
            -ErrorAction Stop | Out-Null
    }
    catch {
        throw "Create-ResourceAccounts failed for $upn : $($_.Exception.Message)"
    }

    return $upn
}

function New-StandardTestUsers {
    param(
        [Parameter(Mandatory)][string]$PSTNGW,
        [string]$DisplayNamePrefix = "Test User",
        [string]$UserNamePrefix    = "testuser",
        [string]$TempPassword      = "P@ssw0rd!ChangeMe123",
        [string]$UsageLocation     = "GB"
    )

    $created = @()

    foreach ($n in 1..2) {
        $upn = "$UserNamePrefix$n@$PSTNGW"
        try {
            New-MgUser -AccountEnabled:$true `
                       -DisplayName "$DisplayNamePrefix $n" `
                       -MailNickname "$UserNamePrefix$n" `
                       -UserPrincipalName $upn `
                       -UsageLocation $UsageLocation `
                       -PasswordProfile @{
                            ForceChangePasswordNextSignIn = $true
                            Password = $TempPassword
                       } -ErrorAction Stop | Out-Null
        }
        catch { }

        $created += $upn
    }

    return $created
}

function Assign-UsageLocation {
    param([string]$UserPrincipalName, [string]$Country)

    try {
        Set-MgUser -UserId $UserPrincipalName -UsageLocation $Country -ErrorAction Stop
        return $true
    }
    catch {
        Write-Warning "Failed to set usage location for ${UserPrincipalName}: $($_.Exception.Message)"
        return $false
    }
}

function Remove-CreatedTestObjects {
    param([string[]]$UserUpns, [string[]]$ResourceUpns)

    $results = @()

    foreach ($upn in $ResourceUpns) {
        try {
            Remove-MgUser -UserId $upn -ErrorAction Stop
            $results += "Removed Entra ID object: $upn"
        }
        catch {
            $results += "AAD delete skipped/failed for $upn : $($_.Exception.Message)"
        }
    }

    foreach ($upn in $UserUpns) {
        try {
            Remove-MgUser -UserId $upn -ErrorAction Stop
            $results += "Removed user: $upn"
        }
        catch {
            $results += "Failed to remove user $upn : $($_.Exception.Message)"
        }
    }

    return $results
}

# ============================================================
# UK Voice Configuration
# ============================================================

function New-UKVoiceConfiguration {
    param(
        [Parameter(Mandatory)][string]$Site,
        [Parameter(Mandatory)][string]$Country,
        [Parameter(Mandatory)][string[]]$PstnGateways,
        [bool]$DerivedTrunkModel = $true
    )

    if (-not $DerivedTrunkModel) {
        foreach ($gw in $PstnGateways) {
            try {
                New-CsOnlinePSTNGateway `
                    -Fqdn $gw `
                    -SipSignalingPort 5067 `
                    -MaxConcurrentSessions 60 `
                    -Enabled $true `
                    -ForwardCallHistory $true `
                    -MediaBypass $true `
                    -ErrorAction Stop | Out-Null
            }
            catch {
                if ($_.Exception.Message -notmatch 'already exists') {
                    throw "Failed to create PSTN Gateway $gw : $($_.Exception.Message)"
                }
            }
        }
    }

    New-CsTenantDialPlan "${Site}-DP" -Description "Dialplan for ${Site}" -ErrorAction SilentlyContinue | Out-Null

    $nRules = @(
        New-CsVoiceNormalizationRule -Name "${Site}-TollFree-NR" -Parent "${Site}-DP" -Pattern '^0((80(0\d{6,7}|8\d{7}|01111)|500\d{6}))\d*$' -Translation '+44$1' -InMemory -Description "TollFree number normalization for United Kingdom",
        New-CsVoiceNormalizationRule -Name "${Site}-Premium-NR"  -Parent "${Site}-DP" -Pattern '^0((9[018]\d|87[123]|70\d)\d{7})$' -Translation '+44$1' -InMemory -Description "Premium number normalization for United Kingdom",
        New-CsVoiceNormalizationRule -Name "${Site}-Mobile-NR"   -Parent "${Site}-DP" -Pattern '^0((7([1-57-9]\d{8}|624\d{6})))$' -Translation '+44$1' -InMemory -Description "Mobile number normalization for United Kingdom",
        New-CsVoiceNormalizationRule -Name "${Site}-National-NR" -Parent "${Site}-DP" -Pattern '^0((1[1-9]\d{7,8}|2[03489]\d{8}|3[0347]\d{8}|5[56]\d{8}|8((4[2-5]|70)\d{7}|45464\d)))\d*(\D+\d+)?$' -Translation '+44$1' -InMemory -Description "National number normalization for United Kingdom",
        New-CsVoiceNormalizationRule -Name "${Site}-Service-NR"  -Parent "${Site}-DP" -Pattern '^(1(47\d|70\d|800\d|1[68]\d{3}|\d\d)|999|[\*\#][\*\#\d]*\#)$' -Translation '$1' -InMemory -Description "Service number normalization for United Kingdom",
        New-CsVoiceNormalizationRule -Name "${Site}-International-NR" -Parent "${Site}-DP" -Pattern '^(?:00)?(1|7|2[07]|3[0-46]|39\d|4[013-9]|5[1-8]|6[0-6]|8[1246]|9[0-58]|2[1235689]\d|24[013-9]|242\d|3[578]\d|42\d|5[09]\d|6[789]\d|8[035789]\d|9[679]\d)(?:0)?(\d{6,14})(\D+\d+)?$' -Translation '+$1$2' -InMemory -Description "International number normalization for United Kingdom"
    )
    Set-CsTenantDialPlan -Identity "${Site}-DP" -NormalizationRules @{add=$nRules} -ErrorAction SilentlyContinue

    Set-CsOnlinePstnUsage -Identity Global -Usage @{Add="${Country}-Mobile-PU"}             -ErrorAction SilentlyContinue
    Set-CsOnlinePstnUsage -Identity Global -Usage @{Add="${Country}-PersonalNumber-PU"}     -ErrorAction SilentlyContinue
    Set-CsOnlinePstnUsage -Identity Global -Usage @{Add="${Country}-FreePhone-PU"}          -ErrorAction SilentlyContinue
    Set-CsOnlinePstnUsage -Identity Global -Usage @{Add="${Country}-Premium-PU"}            -ErrorAction SilentlyContinue
    Set-CsOnlinePstnUsage -Identity Global -Usage @{Add="${Country}-National-PU"}           -ErrorAction SilentlyContinue
    Set-CsOnlinePstnUsage -Identity Global -Usage @{Add="${Country}-SharedCost-PU"}         -ErrorAction SilentlyContinue
    Set-CsOnlinePstnUsage -Identity Global -Usage @{Add="${Country}-Service-PU"}            -ErrorAction SilentlyContinue
    Set-CsOnlinePstnUsage -Identity Global -Usage @{Add="${Country}-DirectoryEnquiries-PU"} -ErrorAction SilentlyContinue
    Set-CsOnlinePstnUsage -Identity Global -Usage @{Add="${Country}-International-PU"}      -ErrorAction SilentlyContinue

    New-CsOnlineVoiceRoute -Name "${Country}-Mobile-VR"             -Priority 1 -OnlinePstnUsages "${Country}-Mobile-PU"             -OnlinePstnGatewayList $PstnGateways -NumberPattern '^\+44(7([1-57-9]\d{8}|624\d{6}))$' -Description "$Country mobile route" -ErrorAction SilentlyContinue
    New-CsOnlineVoiceRoute -Name "${Country}-PersonalNumber-VR"     -Priority 2 -OnlinePstnUsages "${Country}-PersonalNumber-PU"     -OnlinePstnGatewayList $PstnGateways -NumberPattern '^\+4470\d{8}$' -Description "$Country personal number route" -ErrorAction SilentlyContinue
    New-CsOnlineVoiceRoute -Name "${Country}-FreePhone-VR"          -Priority 3 -OnlinePstnUsages "${Country}-FreePhone-PU"          -OnlinePstnGatewayList $PstnGateways -NumberPattern '^\+44(80(0\d{6,7}|8\d{7}|01111)|500\d{6})$' -Description "$Country FreePhone route" -ErrorAction SilentlyContinue
    New-CsOnlineVoiceRoute -Name "${Country}-Premium-VR"            -Priority 4 -OnlinePstnUsages "${Country}-Premium-PU"            -OnlinePstnGatewayList $PstnGateways -NumberPattern '^\+449[018]\d{8}$' -Description "$Country premium route" -ErrorAction SilentlyContinue
    New-CsOnlineVoiceRoute -Name "${Country}-National-VR"           -Priority 5 -OnlinePstnUsages "${Country}-National-PU"           -OnlinePstnGatewayList $PstnGateways -NumberPattern '^\+440?(1[1-9]\d{7,8}|2[03489]\d{8}|3[0347]\d{8}|5[56]\d{8})' -Description "$Country national route" -ErrorAction SilentlyContinue
    New-CsOnlineVoiceRoute -Name "${Country}-SharedCost-VR"         -Priority 6 -OnlinePstnUsages "${Country}-SharedCost-PU"         -OnlinePstnGatewayList $PstnGateways -NumberPattern '^\+44(8((4[3-5]|7[0-3])\d{7}|45464\d))$' -Description "$Country shared cost route" -ErrorAction SilentlyContinue
    New-CsOnlineVoiceRoute -Name "${Country}-Service-VR"            -Priority 7 -OnlinePstnUsages "${Country}-Service-PU"            -OnlinePstnGatewayList $PstnGateways -NumberPattern '^\+?(1(47\d|70\d|800\d|16\d{3}|\d\d)|999)$' -Description "$Country service route" -ErrorAction SilentlyContinue
    New-CsOnlineVoiceRoute -Name "${Country}-DirectoryEnquiries-VR" -Priority 8 -OnlinePstnUsages "${Country}-DirectoryEnquiries-PU" -OnlinePstnGatewayList $PstnGateways -NumberPattern '^(\+44)?118\d{3}$' -Description "$Country directory services route" -ErrorAction SilentlyContinue
    New-CsOnlineVoiceRoute -Name "${Country}-International-VR"      -Priority 9 -OnlinePstnUsages "${Country}-International-PU"      -OnlinePstnGatewayList $PstnGateways -NumberPattern '^\+((1[2-9]\d\d[2-9]\d{6})|((?!(44))([2-9]\d{6,14})))' -Description "$Country international route" -ErrorAction SilentlyContinue

    New-CsOnlineVoiceRoutingPolicy "$Site-National-VP"      -OnlinePstnUsages "$Country-Mobile-PU", "$Country-FreePhone-PU", "$Country-National-PU", "$Country-SharedCost-PU", "$Country-Service-PU" -Description "$Country national voice policy" -ErrorAction SilentlyContinue
    New-CsOnlineVoiceRoutingPolicy "$Site-International-VP" -OnlinePstnUsages "$Country-Mobile-PU", "$Country-FreePhone-PU", "$Country-National-PU", "$Country-SharedCost-PU", "$Country-Service-PU", "$Country-International-PU" -Description "$Country international voice policy" -ErrorAction SilentlyContinue
    New-CsOnlineVoiceRoutingPolicy "$Site-Premium-VP"       -OnlinePstnUsages "$Country-Mobile-PU", "$Country-PersonalNumber-PU", "$Country-FreePhone-PU", "$Country-Premium-PU", "$Country-National-PU", "$Country-SharedCost-PU", "$Country-Service-PU", "$Country-DirectoryEnquiries-PU", "$Country-International-PU" -Description "$Country premium voice policy" -ErrorAction SilentlyContinue
}

# ============================================================
# Exports
# ============================================================

Export-ModuleMember -Function *-TVW*, `
    Test-GraphConnection, Test-TeamsConnection, `
    Get-PstnGatewaysFromGammaEndpoint, Get-PstnGatewaysFromInput, New-GammaDomains, Get-GammaDomainTxtRecords, Test-GammaDomainsVerification, `
    Normalize-SkuPartNumber, Get-CommercialTargetSkus, Get-LicenseInventory, Assign-LicenseSku, `
    Create-ResourceAccounts, New-StandardTestUsers, Assign-UsageLocation, Remove-CreatedTestObjects, `
    New-UKVoiceConfiguration