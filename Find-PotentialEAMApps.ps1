#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Identifies potential Enterprise App Management (EAM) catalog apps from Intune detected software.

.DESCRIPTION
    Interactive (or delegated/app) Graph script that:
      1. Loads the EAM catalog packages available in the tenant
      2. Loads win32CatalogApp apps already managed in Intune
      3. Loads Intune detected apps (display name, version, deviceCount)
      4. Matches detected software to catalog products
      5. Prioritizes candidates that appear on multiple devices and show version
         fragmentation or lag behind the catalog (likely not self-updating well)
      6. Optionally creates unassigned (or group-assigned) EAM apps from the catalog

    This is a discovery / onboarding helper. It does not replace Invoke-EAMAutoUpdate.ps1
    (which updates apps already under management).

.PARAMETER MinDeviceCount
    Minimum total devices for a catalog match to appear as a candidate. Default: 5.

.PARAMETER MinDistinctVersions
    Optional filter: only show products with at least this many distinct detected versions.
    Default: 1 (no extra filter). Raise (e.g. 2) to focus on fragmented installs.

.PARAMETER Top
    Maximum candidates to return after sorting. Default: 50. Use 0 for no limit.

.PARAMETER ExportPath
    Optional CSV path for the full candidate report.

.PARAMETER AddApps
    After reporting, add selected candidates as EAM (win32CatalogApp) apps in Intune.
    Without -AddAppNames and without -AddAll, you are prompted per candidate (Y/n).

.PARAMETER AddAll
    With -AddApps: add all candidates that pass filters (no per-app prompt). Implies care.

.PARAMETER AddAppNames
    With -AddApps: only add candidates whose CatalogProductName matches one of these names
    (case-insensitive). Skips interactive prompts for other apps.

.PARAMETER AssignGroupId
    Optional Entra group object ID. When set, new apps get an "available" assignment to this group.
    When omitted, apps are created unassigned.

.PARAMETER AssignmentNotifications
    End-user toast notification mode for new group assignments (when -AssignGroupId is set).
    Valid values: hideAll (default), showReboot, showAll.
    Matches Microsoft Graph win32LobAppNotification. Subsequent updates via Invoke-EAMAutoUpdate
    preserve whatever is configured on the previous assignment.

.PARAMETER AssignmentDeliveryOptimization
    Content download priority for new group assignments (when -AssignGroupId is set).
    Valid values: notConfigured (default — download in background), foreground.
    Matches Microsoft Graph win32LobAppDeliveryOptimizationPriority. Subsequent updates via
    Invoke-EAMAutoUpdate preserve whatever is configured on the previous assignment.

.PARAMETER ExcludeCatalogNames
    Catalog product names to ignore (e.g. already rejected titles).

.PARAMETER IncludeAlreadyManaged
    Include catalog products that already have a win32CatalogApp in the tenant.
    Default: only products not yet managed via EAM.

.PARAMETER WhatIfAdd
    Show which apps would be created without calling the create APIs.

.EXAMPLE
    # Discover and export candidates (read-only)
    .\Find-PotentialEAMApps.ps1 -MinDeviceCount 10 -ExportPath .\eam-candidates.csv

.EXAMPLE
    # Focus on multi-version (poor self-update) installs and prompt to add
    .\Find-PotentialEAMApps.ps1 -MinDeviceCount 5 -MinDistinctVersions 2 -AddApps

.EXAMPLE
    # Add specific catalog titles after discovery (assignments default to hideAll toasts)
    .\Find-PotentialEAMApps.ps1 -AddApps -AddAppNames 'Notepad++','7-Zip' -AssignGroupId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'

.EXAMPLE
    # Assign with full end-user notifications instead of the hideAll default
    .\Find-PotentialEAMApps.ps1 -AddApps -AddAppNames '7-Zip' -AssignGroupId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' -AssignmentNotifications showAll

.EXAMPLE
    # Assign with foreground content download instead of background (notConfigured)
    .\Find-PotentialEAMApps.ps1 -AddApps -AddAppNames '7-Zip' -AssignGroupId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' -AssignmentDeliveryOptimization foreground
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100000)]
    [int]
    $MinDeviceCount = 5,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 1000)]
    [int]
    $MinDistinctVersions = 1,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 100000)]
    [int]
    $Top = 50,

    [Parameter(Mandatory = $false)]
    [string]
    $ExportPath,

    [Parameter(Mandatory = $false)]
    [switch]
    $AddApps,

    [Parameter(Mandatory = $false)]
    [switch]
    $AddAll,

    [Parameter(Mandatory = $false)]
    [string[]]
    $AddAppNames = @(),

    [Parameter(Mandatory = $false)]
    [string]
    $AssignGroupId,

    [Parameter(Mandatory = $false)]
    [ValidateSet('hideAll', 'showReboot', 'showAll')]
    [string]
    $AssignmentNotifications = 'hideAll',

    [Parameter(Mandatory = $false)]
    [ValidateSet('notConfigured', 'foreground')]
    [string]
    $AssignmentDeliveryOptimization = 'notConfigured',

    [Parameter(Mandatory = $false)]
    [string[]]
    $ExcludeCatalogNames = @(),

    [Parameter(Mandatory = $false)]
    [switch]
    $IncludeAlreadyManaged,

    [Parameter(Mandatory = $false)]
    [switch]
    $WhatIfAdd
)

$ErrorActionPreference = 'Stop'

#region Helpers

function Connect-EAMDiscoveryGraph {
    $scopes = @(
        'DeviceManagementApps.Read.All'
        'DeviceManagementManagedDevices.Read.All'
    )
    if ($AddApps) {
        $scopes += 'DeviceManagementApps.ReadWrite.All'
    }

    $context = Get-MgContext -ErrorAction SilentlyContinue
    $haveScopes = $false
    if ($context -and $context.Scopes) {
        $missing = $scopes | Where-Object { $context.Scopes -notcontains $_ }
        $haveScopes = ($missing.Count -eq 0)
    }

    if (-not $haveScopes) {
        Write-Host "Connecting to Microsoft Graph with scopes: $($scopes -join ', ')"
        Connect-MgGraph -Scopes $scopes -NoWelcome | Out-Null
    }
    else {
        Write-Host "Using existing Graph context for tenant $($context.TenantId)."
    }

    $context = Get-MgContext
    if (-not $context) {
        throw 'Failed to establish a Microsoft Graph context.'
    }

    Write-Host "Connected. TenantId=$($context.TenantId); Account=$($context.Account); AuthType=$($context.AuthType)"
}

function Invoke-GraphGetAll {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $Uri
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri

    while ($next) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        foreach ($value in @($page.value)) {
            $items.Add($value)
        }
        $next = $page.'@odata.nextLink'
    }

    return $items
}

# Tokens too common for reliable catalog matching (cause false positives like SQL CU / Runtime).
# Also generic product words that alone must never link unrelated apps (Authenticator, Driver, …).
$script:MatchStopwords = [System.Collections.Generic.HashSet[string]]::new([string[]]@(
        'the', 'for', 'and', 'app', 'application', 'applications', 'desktop', 'windows', 'microsoft',
        'client', 'service', 'services', 'runtime', 'update', 'updates', 'cumulative', 'latest',
        'plugin', 'tools', 'toolkit', 'agent', 'machine', 'connected', 'webview', 'webview2',
        'server', 'sql', 'with', 'from', 'version', 'package', 'installer', 'setup', 'x64', 'x86',
        'bit', 'amd64', 'arm64', 'en', 'us', 'business', 'enterprise', 'professional', 'standard',
        'redistributable', 'redist', 'host', 'hosting', 'bundle', 'component', 'components',
        'visual', 'studio', 'team', 'explorer', 'office', 'system', 'framework', 'core', 'sdk',
        'driver', 'drivers', 'authenticator', 'manager', 'helper', 'utility', 'utilities', 'tool',
        'viewer', 'player', 'connector', 'provider', 'library', 'libraries', 'shared', 'common',
        'audio', 'video', 'codec', 'control', 'panel', 'software', 'suite', 'pack', 'kit', 'bridge',
        'support', 'assistant', 'integrity', 'evidence', 'ole', 'db', 'odbc', 'jdbc', 'native',
        'redistributable', 'launcher', 'bootstrapper', 'prerequisites', 'prerequisite'
    ), [StringComparer]::OrdinalIgnoreCase)

# Channel / ring SKUs that identify different products (Edge Dev != Edge, Chrome Beta != Chrome).
# These are NOT stopwords — they must be preserved and enforced during matching.
$script:ChannelIdentityTokens = [System.Collections.Generic.HashSet[string]]::new([string[]]@(
        'dev', 'beta', 'canary', 'nightly', 'preview', 'insider', 'insiders',
        'ltsc', 'esr', 'unstable', 'stable', 'sxs', 'development', 'developer'
    ), [StringComparer]::OrdinalIgnoreCase)

function Test-IsProductYearToken {
    param ([string]$Token)
    # Product-line years (VS 2022 vs 2024), not dotted package versions.
    return ($Token -match '^(19|20)\d{2}$')
}

function Test-IsProductMinorToken {
    param ([string]$Token)
    # Encoded major.minor product lines: Python 3.11 -> 3_11 (not full package 3.11.9.x builds).
    return ($Token -match '^\d+_\d+$')
}

function Test-IsChannelIdentityToken {
    param ([string]$Token)
    if ([string]::IsNullOrWhiteSpace($Token)) {
        return $false
    }
    return $script:ChannelIdentityTokens.Contains($Token)
}

function Get-ChannelIdentityTokens {
    param ([string[]]$Tokens)
    return @($Tokens | Where-Object { Test-IsChannelIdentityToken -Token $_ })
}

function Test-TokenSetsIntersect {
    param (
        [string[]]$Left,
        [string[]]$Right
    )
    foreach ($item in $Left) {
        if ($Right -contains $item) {
            return $true
        }
    }
    return $false
}

function Test-ProductIdentityCompatible {
    <#
    .SYNOPSIS
    Returns $false when product-line identity disagrees:
    - years (VS 2022 vs 2024)
    - channels (Edge Dev vs Edge)
    - major.minor runtimes (Python 3.11 vs 3.12)
    #>
    param (
        [string[]]$DetectedTokens,
        [string[]]$CatalogTokens
    )

    $detectedYears = @($DetectedTokens | Where-Object { Test-IsProductYearToken -Token $_ })
    $catalogYears = @($CatalogTokens | Where-Object { Test-IsProductYearToken -Token $_ })
    if ($detectedYears.Count -gt 0 -and $catalogYears.Count -gt 0) {
        if (-not (Test-TokenSetsIntersect -Left $detectedYears -Right $catalogYears)) {
            return $false
        }
    }

    $detectedMinors = @($DetectedTokens | Where-Object { Test-IsProductMinorToken -Token $_ })
    $catalogMinors = @($CatalogTokens | Where-Object { Test-IsProductMinorToken -Token $_ })
    # Python 3.11 must not map to Python 3.12; generic "Python" must not absorb a minor-specific catalog entry.
    if ($detectedMinors.Count -gt 0 -and $catalogMinors.Count -eq 0) {
        return $false
    }
    if ($catalogMinors.Count -gt 0 -and $detectedMinors.Count -eq 0) {
        return $false
    }
    if ($detectedMinors.Count -gt 0 -and $catalogMinors.Count -gt 0) {
        if (-not (Test-TokenSetsIntersect -Left $detectedMinors -Right $catalogMinors)) {
            return $false
        }
    }

    $detectedChannels = Get-ChannelIdentityTokens -Tokens $DetectedTokens
    $catalogChannels = Get-ChannelIdentityTokens -Tokens $CatalogTokens

    # Stable (no channel token) must not match a channel SKU, and vice versa.
    if ($detectedChannels.Count -gt 0 -and $catalogChannels.Count -eq 0) {
        return $false
    }
    if ($catalogChannels.Count -gt 0 -and $detectedChannels.Count -eq 0) {
        return $false
    }
    if ($detectedChannels.Count -gt 0 -and $catalogChannels.Count -gt 0) {
        if (-not (Test-TokenSetsIntersect -Left $detectedChannels -Right $catalogChannels)) {
            return $false
        }
    }

    return $true
}

function Test-IsMatchToken {
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $Token
    )

    if ([string]::IsNullOrWhiteSpace($Token)) {
        return $false
    }

    # Keep product years as identity (Community 2022 != Community 2024).
    if (Test-IsProductYearToken -Token $Token) {
        return $true
    }

    # Keep major.minor product lines (Python 3_11 != 3_12).
    if (Test-IsProductMinorToken -Token $Token) {
        return $true
    }

    # Channel tokens (dev/beta/canary) are identity even when shorter than 4 chars ("dev").
    if (Test-IsChannelIdentityToken -Token $Token) {
        return $true
    }

    if ($Token.Length -lt 4) {
        return $false
    }
    # Other pure numbers (build ids) are not useful match keys.
    if ($Token -match '^\d+$') {
        return $false
    }
    if ($script:MatchStopwords.Contains($Token)) {
        return $false
    }
    return $true
}

function Get-MatchTokens {
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $NormalizedName
    )

    if ([string]::IsNullOrWhiteSpace($NormalizedName)) {
        return @()
    }

    return @(
        $NormalizedName -split '\s+' |
            Where-Object { Test-IsMatchToken -Token $_ }
    )
}

function Get-NormalizedAppName {
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $Name,

        [Parameter(Mandatory = $false)]
        [switch]
        $KeepStopwords
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ''
    }

    $n = $Name.ToLowerInvariant()
    # Preserve meaningful tech tokens before stripping punctuation.
    $n = $n -replace 'c\+\+', 'cplusplus'
    $n = $n -replace '\.net', 'dotnet'
    # Strip common bitness / architecture / installer noise for matching.
    $n = $n -replace '\(x64\)|\(x86\)|\(64-bit\)|\(32-bit\)|64-bit|32-bit|x64|x86', ' '
    $n = $n -replace '\b(version|ver|v)\s*[\d\._]+', ' '

    # Encode runtime/language major.minor product lines before stripping builds:
    #   Python 3.11 / 3.11.9  -> 3_11
    #   .NET Runtime 6.0.36   -> 6_0   (distinct from 7_0)
    #   .NET Runtime 7.0      -> 7_0
    # Majors 1-99 only so browser builds (Chrome/Edge 100+) are not treated as product lines.
    $n = $n -replace '\b([1-9]\d?)\.(\d{1,2})(?:\.\d+)*\b', '$1_$2'

    # Remaining multi-part dotted builds (Chrome 150.0.7871.115, etc.).
    $n = $n -replace '\d+(?:\.\d+){2,}', ' '
    $n = $n -replace '[^a-z0-9_]+', ' '
    $tokens = @(
        $n -split '\s+' | Where-Object {
            if (-not $_) { return $false }
            if ($KeepStopwords) {
                if (Test-IsProductYearToken -Token $_) { return $true }
                if (Test-IsProductMinorToken -Token $_) { return $true }
                if (Test-IsChannelIdentityToken -Token $_) { return $true }
                return ($_ -notmatch '^\d+$')
            }
            return (Test-IsMatchToken -Token $_)
        }
    )
    return ($tokens -join ' ').Trim()
}

function Get-GraphPropertyValue {
    param (
        [Parameter(Mandatory = $true)]
        $Object,

        [Parameter(Mandatory = $true)]
        [string[]]
        $Names
    )

    foreach ($name in $Names) {
        if ($null -eq $Object) {
            continue
        }

        $prop = $Object.PSObject.Properties[$name]
        if ($prop -and $null -ne $prop.Value -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
            return $prop.Value
        }

        if ($Object -is [System.Collections.IDictionary] -and $Object.Contains($name) -and $Object[$name]) {
            return $Object[$name]
        }

        if ($Object.AdditionalProperties -is [System.Collections.IDictionary] -and
            $Object.AdditionalProperties.ContainsKey($name) -and
            $Object.AdditionalProperties[$name]) {
            return $Object.AdditionalProperties[$name]
        }
    }

    return $null
}

function ConvertTo-ComparableVersion {
    <#
    .SYNOPSIS
    Fast, bounded version parse for inventory strings (avoids SemanticVersion edge-case cost).
    .OUTPUTS
    System.Version or $null
    #>
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $VersionString
    )

    if ([string]::IsNullOrWhiteSpace($VersionString)) {
        return $null
    }

    # Inventory sometimes returns huge/garbage strings; never spend time parsing those.
    $normalized = $VersionString.Trim()
    if ($normalized.Length -gt 64 -or $normalized -eq 'unknown') {
        return $null
    }

    $core = ($normalized -split '[\s\+]', 2)[0]
    $core = ($core -split '-', 2)[0]
    $core = $core -replace '[^\d\.]', ''
    if ([string]::IsNullOrWhiteSpace($core) -or $core -notmatch '\d') {
        return $null
    }

    $parts = [System.Collections.Generic.List[int]]::new()
    foreach ($segment in $core.Split('.', [System.StringSplitOptions]::RemoveEmptyEntries)) {
        if ($segment.Length -gt 9) {
            # Avoid OverflowException / pathological segments
            return $null
        }
        $number = 0
        if (-not [int]::TryParse($segment, [ref]$number)) {
            return $null
        }
        $parts.Add($number)
        if ($parts.Count -ge 4) {
            break
        }
    }

    if ($parts.Count -eq 0) {
        return $null
    }

    while ($parts.Count -lt 4) {
        $parts.Add(0)
    }

    try {
        return [version]::new($parts[0], $parts[1], $parts[2], $parts[3])
    }
    catch {
        return $null
    }
}

function Compare-VersionString {
    param (
        [string]$Left,
        [string]$Right
    )

    if ([string]::IsNullOrWhiteSpace($Left) -or [string]::IsNullOrWhiteSpace($Right)) {
        return $null
    }
    if ($Left.Trim() -eq $Right.Trim()) {
        return 0
    }

    $l = ConvertTo-ComparableVersion -VersionString $Left
    $r = ConvertTo-ComparableVersion -VersionString $Right
    if ($null -eq $l -or $null -eq $r) {
        return $null
    }
    return $l.CompareTo($r)
}

function Get-CatalogPackageVersion {
    param ($Package)

    # Graph mobileAppCatalogPackage uses versionDisplayName (see sample properties from tenant).
    return [string](Get-GraphPropertyValue -Object $Package -Names @(
            'versionDisplayName'
            'version'
            'packageVersion'
            'displayVersion'
            'packageFullVersion'
            'packageVersionDisplayName'
            'branchVersion'
            'productVersion'
        ))
}

function Get-EAMCatalogProducts {
    Write-Host 'Loading Enterprise App Catalog packages...'
    $packages = Invoke-GraphGetAll -Uri 'https://graph.microsoft.com/beta/deviceAppManagement/mobileAppCatalogPackages?$top=100'

    if ($packages.Count -eq 0) {
        throw 'No mobileAppCatalogPackages returned. Confirm the tenant has Enterprise App Management / catalog access.'
    }

    Write-Host "  Retrieved $($packages.Count) catalog package revision(s)."

    # One-time diagnostic of package shape (helps when version/id mapping drifts).
    $sample = $packages[0]
    $sampleKeys = @($sample.PSObject.Properties.Name)
    Write-Host "  Sample package properties: $($sampleKeys -join ', ')"

    # Group by product; keep latest package id by version comparison when possible.
    # Note: do not use $PID — that is a read-only automatic variable (process id).
    $byProduct = $packages | Group-Object -Property {
        $catalogProductId = Get-GraphPropertyValue -Object $_ -Names @('productId', 'productIdentifier')
        if ($catalogProductId) { [string]$catalogProductId }
        else {
            $name = Get-GraphPropertyValue -Object $_ -Names @('productName', 'productDisplayName', 'displayName', 'title', 'name')
            if ($name) { [string]$name } else { [string](Get-GraphPropertyValue -Object $_ -Names @('id')) }
        }
    }

    $products = [System.Collections.Generic.List[object]]::new()
    $missingVersion = 0
    $missingPackageId = 0

    foreach ($group in $byProduct) {
        $revisions = @($group.Group)
        $productName = $null
        $publisher = $null
        $productId = $null

        foreach ($rev in $revisions) {
            if (-not $productName) {
                $productName = Get-GraphPropertyValue -Object $rev -Names @(
                    'productDisplayName', 'productName', 'displayName', 'title', 'name', 'packageDisplayName'
                )
            }
            if (-not $publisher) {
                $publisher = Get-GraphPropertyValue -Object $rev -Names @(
                    'publisherDisplayName', 'publisher', 'publisherName'
                )
            }
            if (-not $productId) {
                $productId = Get-GraphPropertyValue -Object $rev -Names @('productId', 'productIdentifier')
            }
        }

        if ([string]::IsNullOrWhiteSpace([string]$productName)) {
            continue
        }

        $best = $revisions[0]
        $bestVersion = Get-CatalogPackageVersion -Package $best
        foreach ($rev in $revisions) {
            $revVersion = Get-CatalogPackageVersion -Package $rev
            $cmp = Compare-VersionString -Left ([string]$revVersion) -Right ([string]$bestVersion)
            if ($cmp -eq 1) {
                $best = $rev
                $bestVersion = $revVersion
            }
        }

        # If no comparable versions, prefer the last revision in the page group (often newest).
        if ([string]::IsNullOrWhiteSpace([string]$bestVersion) -and $revisions.Count -gt 1) {
            $best = $revisions[-1]
            $bestVersion = Get-CatalogPackageVersion -Package $best
        }

        $packageId = Get-GraphPropertyValue -Object $best -Names @(
            'id', 'packageId', 'mobileAppCatalogPackageId', 'catalogPackageId'
        )

        if ([string]::IsNullOrWhiteSpace([string]$bestVersion)) {
            $missingVersion++
        }
        if ([string]::IsNullOrWhiteSpace([string]$packageId)) {
            $missingPackageId++
        }

        $normalizedStrict = Get-NormalizedAppName -Name ([string]$productName)
        $normalizedSoft = Get-NormalizedAppName -Name ([string]$productName) -KeepStopwords
        # Prefer strict tokens for matching; fall back to soft when stopwords would wipe the name.
        $normalized = if ($normalizedStrict) { $normalizedStrict } else { $normalizedSoft }

        $products.Add([PSCustomObject]@{
                ProductId       = [string]$productId
                ProductName     = [string]$productName
                Publisher       = [string]$publisher
                LatestVersion   = [string]$bestVersion
                LatestPackageId = [string]$packageId
                NormalizedName  = $normalized
                RevisionCount   = $revisions.Count
            })
    }

    Write-Host "  Catalog products (unique): $($products.Count)"
    if ($missingVersion -gt 0) {
        Write-Warning "  $missingVersion product(s) have no resolvable catalog version property."
    }
    if ($missingPackageId -gt 0) {
        Write-Warning "  $missingPackageId product(s) have no resolvable package id (cannot -AddApps those)."
    }
    $withVersion = @($products | Where-Object { $_.LatestVersion }).Count
    Write-Host "  Products with catalog version: $withVersion / $($products.Count)"

    return $products
}

function Get-ManagedWin32CatalogApps {
    Write-Host 'Loading existing Intune win32CatalogApp (EAM) apps...'
    # Avoid $select so derived win32CatalogApp fields (if present) are returned.
    $filter = [uri]::EscapeDataString("isof('microsoft.graph.win32CatalogApp')")
    $uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=$filter&`$top=100"
    $apps = Invoke-GraphGetAll -Uri $uri
    Write-Host "  Managed EAM apps: $($apps.Count)"

    return @(
        $apps | ForEach-Object {
            $displayName = [string](Get-GraphPropertyValue -Object $_ -Names @('displayName', 'name'))
            $normalizedStrict = Get-NormalizedAppName -Name $displayName
            $normalizedSoft = Get-NormalizedAppName -Name $displayName -KeepStopwords
            $normalized = if ($normalizedStrict) { $normalizedStrict } else { $normalizedSoft }

            [PSCustomObject]@{
                Id             = [string](Get-GraphPropertyValue -Object $_ -Names @('id'))
                DisplayName    = $displayName
                Publisher      = [string](Get-GraphPropertyValue -Object $_ -Names @('publisher', 'publisherDisplayName'))
                DisplayVersion = [string](Get-GraphPropertyValue -Object $_ -Names @('displayVersion', 'version'))
                ProductId      = [string](Get-GraphPropertyValue -Object $_ -Names @('productId', 'mobileAppCatalogPackageId', 'packageId'))
                NormalizedName = $normalized
            }
        }
    )
}

function Get-CatalogProductKey {
    param (
        [Parameter(Mandatory = $true)]
        $CatalogProduct
    )

    if ($CatalogProduct.ProductId) {
        return "id:$($CatalogProduct.ProductId)"
    }
    return "name:$($CatalogProduct.ProductName)"
}

function Get-AliasNormalizedAppName {
    <#
    .SYNOPSIS
    Looser name form for matching Intune EAM display names to catalog titles.
    Strips architecture and "for Windows/Business" style suffixes that often differ
    between portal naming and catalog productDisplayName.
    #>
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ''
    }

    $n = $Name.ToLowerInvariant()
    $n = $n -replace 'c\+\+', 'cplusplus'
    $n = $n -replace '\.net', 'dotnet'
    $n = $n -replace '\(x64\)|\(x86\)|\(64-bit\)|\(32-bit\)|64-bit|32-bit|\bx64\b|\bx86\b', ' '
    # Tenant naming often differs: "Edge for Windows" vs catalog "Edge for Business"
    $n = $n -replace '\bfor\s+(windows|business|enterprise|education|workstations)\b', ' '
    $n = $n -replace '\b(version|ver|v)\s*[\d\._]+', ' '
    # Keep runtime major.minor product lines; drop long dotted builds.
    $n = $n -replace '\b([1-9]\d?)\.(\d{1,2})(?:\.\d+)*\b', '$1_$2'
    $n = $n -replace '\d+(?:\.\d+){2,}', ' '
    $n = $n -replace '[^a-z0-9_]+', ' '

    $weak = [System.Collections.Generic.HashSet[string]]::new([string[]]@(
            'the', 'and', 'for', 'of', 'a', 'to', 'with', 'from', 'by', 'on', 'in'
        ), [StringComparer]::OrdinalIgnoreCase)

    $tokens = @(
        $n -split '\s+' | Where-Object {
            $_ -and -not $weak.Contains($_) -and ($_ -notmatch '^\d+$' -or (Test-IsProductYearToken -Token $_) -or (Test-IsProductMinorToken -Token $_))
        }
    )
    return ($tokens -join ' ').Trim()
}

function Test-AliasAppNamesMatch {
    <#
    .SYNOPSIS
    Returns $true when an Intune EAM app name and a catalog product name refer to the same app
    despite different display naming (Chrome for Business 64-bit vs Google Chrome for Business).
    Intentionally looser than inventory->catalog discovery matching.
    #>
    param (
        [string]$NameA,
        [string]$NameB
    )

    if ([string]::IsNullOrWhiteSpace($NameA) -or [string]::IsNullOrWhiteSpace($NameB)) {
        return $false
    }

    $a = Get-AliasNormalizedAppName -Name $NameA
    $b = Get-AliasNormalizedAppName -Name $NameB
    if (-not $a -or -not $b) {
        return $false
    }

    if ($a -eq $b) {
        return $true
    }

    $tokensA = @($a -split '\s+' | Where-Object { $_ })
    $tokensB = @($b -split '\s+' | Where-Object { $_ })
    if ($tokensA.Count -eq 0 -or $tokensB.Count -eq 0) {
        return $false
    }

    # Identity gates (years / channels / python|dotnet minors) still apply.
    if (-not (Test-ProductIdentityCompatible -DetectedTokens $tokensA -CatalogTokens $tokensB)) {
        return $false
    }

    $setA = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($t in $tokensA) { [void]$setA.Add($t) }
    $setB = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($t in $tokensB) { [void]$setB.Add($t) }

    $shorter = if ($tokensA.Count -le $tokensB.Count) { $tokensA } else { $tokensB }
    $longerSet = if ($tokensA.Count -le $tokensB.Count) { $setB } else { $setA }
    $longer = if ($tokensA.Count -le $tokensB.Count) { $tokensB } else { $tokensA }

    # All tokens from the shorter name must appear in the longer name.
    foreach ($t in $shorter) {
        if (-not $longerSet.Contains($t)) {
            return $false
        }
    }

    # Single-token alias only for distinctive brands (chrome, zoom, …) — not "edge" alone if you want,
    # but edge is 4 chars and is the real product stem after stripping for windows/business.
    if ($shorter.Count -eq 1) {
        $tok = $shorter[0]
        if ($tok.Length -lt 4) {
            return $false
        }
        # Longer side should not be hugely more specific without shared brand prefix context.
        if ($longer.Count -gt 3) {
            return $false
        }
        return $true
    }

    # Multi-token: require solid coverage of the longer name as well.
    $coverage = $shorter.Count / [double]$longer.Count
    return ($coverage -ge 0.5)
}

function Add-ManagedCatalogLookupEntry {
    param (
        [hashtable]$Lookup,
        $CatalogProduct,
        $ManagedApp
    )

    if (-not $CatalogProduct -or -not $ManagedApp) {
        return
    }

    $key = Get-CatalogProductKey -CatalogProduct $CatalogProduct
    if ($Lookup.ContainsKey($key)) {
        return
    }

    $Lookup[$key] = [PSCustomObject]@{
        CatalogProduct = $CatalogProduct
        ManagedAppId   = $ManagedApp.Id
        ManagedAppName = $ManagedApp.DisplayName
        ManagedVersion = $ManagedApp.DisplayVersion
    }
}

function Build-ManagedCatalogLookup {
    <#
    .SYNOPSIS
    Maps catalog products that are already present as win32CatalogApp in Intune.

    Matching is intentionally looser than inventory discovery because tenants often rename
    EAM apps (e.g. "Chrome for Business 64-bit" vs catalog "Google Chrome for Business",
    "Microsoft Edge for Windows" vs "Microsoft Edge for Business").
    #>
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]
        $ManagedApps,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]
        $CatalogProducts,

        [Parameter(Mandatory = $true)]
        $CatalogIndex
    )

    $lookup = @{} # catalog key -> managed app info
    $catalogByProductId = @{}
    $catalogByDisplayName = @{} # lower invariant -> product

    foreach ($product in $CatalogProducts) {
        if ($product.ProductId) {
            $catalogByProductId[[string]$product.ProductId] = $product
        }
        if ($product.ProductName) {
            $catalogByDisplayName[$product.ProductName.ToLowerInvariant()] = $product
        }
    }

    foreach ($managed in $ManagedApps) {
        $matchedProducts = [System.Collections.Generic.List[object]]::new()

        # 1) productId / package id on the managed app object
        if ($managed.ProductId) {
            if ($catalogByProductId.ContainsKey([string]$managed.ProductId)) {
                $matchedProducts.Add($catalogByProductId[[string]$managed.ProductId])
            }
            else {
                foreach ($product in $CatalogProducts) {
                    if ($product.LatestPackageId -eq $managed.ProductId) {
                        $matchedProducts.Add($product)
                    }
                }
            }
        }

        # 2) Exact display name (case-insensitive)
        if ($managed.DisplayName) {
            $nameKey = $managed.DisplayName.ToLowerInvariant()
            if ($catalogByDisplayName.ContainsKey($nameKey)) {
                $matchedProducts.Add($catalogByDisplayName[$nameKey])
            }
        }

        # 3) Exact discovery-normalized name: mark *all* catalog products sharing that key
        if ($managed.NormalizedName -and $CatalogIndex.Exact.ContainsKey($managed.NormalizedName)) {
            foreach ($product in $CatalogIndex.Exact[$managed.NormalizedName]) {
                $matchedProducts.Add($product)
            }
        }

        # 4) Strict fuzzy matcher (multi-token)
        if ($managed.NormalizedName) {
            $fuzzy = Find-CatalogMatchFromIndex -NormalizedName $managed.NormalizedName -Index $CatalogIndex
            if ($fuzzy) {
                $matchedProducts.Add($fuzzy)
            }
        }

        # 5) Alias matcher: handles "Chrome for Business 64-bit" ↔ "Google Chrome for Business"
        #    and "Microsoft Edge for Windows" ↔ "Microsoft Edge for Business"
        foreach ($product in $CatalogProducts) {
            if (Test-AliasAppNamesMatch -NameA $managed.DisplayName -NameB $product.ProductName) {
                $matchedProducts.Add($product)
            }
        }

        foreach ($product in $matchedProducts) {
            Add-ManagedCatalogLookupEntry -Lookup $lookup -CatalogProduct $product -ManagedApp $managed
        }
    }

    Write-Host "  Catalog products already managed as EAM: $($lookup.Count) (from $($ManagedApps.Count) Intune win32CatalogApp apps)"
    if ($lookup.Count -gt 0 -and $lookup.Count -le 50) {
        foreach ($item in ($lookup.Values | Sort-Object ManagedAppName, { $_.CatalogProduct.ProductName })) {
            Write-Host ("    - Intune '{0}' => catalog '{1}'" -f $item.ManagedAppName, $item.CatalogProduct.ProductName)
        }
    }
    elseif ($lookup.Count -gt 50) {
        Write-Host '    (mapping list omitted; more than 50 links)'
    }

    $unmapped = @($ManagedApps | Where-Object {
            $id = $_.Id
            -not ($lookup.Values | Where-Object { $_.ManagedAppId -eq $id } | Select-Object -First 1)
        })
    if ($unmapped.Count -gt 0) {
        Write-Host "  Unmapped Intune EAM apps (no catalog link): $($unmapped.Count)"
        foreach ($u in ($unmapped | Select-Object -First 20)) {
            Write-Host ("    - '{0}' (alias='{1}')" -f $u.DisplayName, (Get-AliasNormalizedAppName -Name $u.DisplayName))
        }
    }

    return $lookup
}

function Test-CatalogProductAlreadyManaged {
    param (
        [Parameter(Mandatory = $true)]
        $CatalogProduct,

        [Parameter(Mandatory = $true)]
        [hashtable]
        $ManagedCatalogLookup
    )

    $key = Get-CatalogProductKey -CatalogProduct $CatalogProduct
    if ($ManagedCatalogLookup.ContainsKey($key)) {
        return $ManagedCatalogLookup[$key]
    }

    # Fallback: case-insensitive product name vs any mapped catalog product name
    foreach ($item in $ManagedCatalogLookup.Values) {
        if ($item.CatalogProduct.ProductName -and $CatalogProduct.ProductName -and
            $item.CatalogProduct.ProductName.Equals($CatalogProduct.ProductName, [StringComparison]::OrdinalIgnoreCase)) {
            return $item
        }
        if ($item.ManagedAppName -and $CatalogProduct.ProductName -and
            $item.ManagedAppName.Equals($CatalogProduct.ProductName, [StringComparison]::OrdinalIgnoreCase)) {
            return $item
        }
    }

    return $null
}

function Get-IntuneDetectedApps {
    Write-Host 'Loading Intune detected apps (device inventory)...'
    # Beta tends to return richer inventory; falls back if needed.
    $uri = 'https://graph.microsoft.com/beta/deviceManagement/detectedApps?$select=id,displayName,version,publisher,deviceCount,sizeInByte&$top=100'
    try {
        $apps = Invoke-GraphGetAll -Uri $uri
    }
    catch {
        Write-Warning "Beta detectedApps failed ($($_.Exception.Message)); trying v1.0..."
        $uri = 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?$select=id,displayName,version,publisher,deviceCount,sizeInByte&$top=100'
        $apps = Invoke-GraphGetAll -Uri $uri
    }

    Write-Host "  Detected app rows: $($apps.Count)"
    return $apps
}

function Select-BestCatalogProduct {
    param (
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]
        $Products
    )

    # Prefer the *shortest* display name among exact normalized collisions
    # (longest previously favored over-broad titles).
    $best = $null
    $bestLength = [int]::MaxValue
    foreach ($p in $Products) {
        $len = ([string]$p.ProductName).Length
        if ($len -gt 0 -and $len -lt $bestLength) {
            $best = $p
            $bestLength = $len
        }
    }
    return $best
}

function New-CatalogMatchIndex {
    <#
    .SYNOPSIS
    Builds O(1) exact and token indexes so matching is not O(detected x catalog).
    #>
    param (
        [Parameter(Mandatory = $true)]
        [object[]]
        $CatalogProducts
    )

    $exact = @{}   # normalizedName -> List[product]
    $byToken = @{} # token -> List[product]

    foreach ($product in $CatalogProducts) {
        $normalized = [string]$product.NormalizedName
        if ([string]::IsNullOrWhiteSpace($normalized)) {
            continue
        }

        if (-not $exact.ContainsKey($normalized)) {
            $exact[$normalized] = [System.Collections.Generic.List[object]]::new()
        }
        $exact[$normalized].Add($product)

        foreach ($token in (Get-MatchTokens -NormalizedName $normalized)) {
            if (-not $byToken.ContainsKey($token)) {
                $byToken[$token] = [System.Collections.Generic.List[object]]::new()
            }
            $byToken[$token].Add($product)
        }
    }

    return [PSCustomObject]@{
        Exact   = $exact
        ByToken = $byToken
    }
}

function Find-CatalogMatchFromIndex {
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]
        $NormalizedName,

        [Parameter(Mandatory = $true)]
        [object]
        $Index
    )

    if ([string]::IsNullOrWhiteSpace($NormalizedName)) {
        return $null
    }

    # 1) Exact normalized-name hit only path that allows weak/generic single words.
    if ($Index.Exact.ContainsKey($NormalizedName)) {
        $exactList = $Index.Exact[$NormalizedName]
        if ($exactList.Count -eq 1) {
            return $exactList[0]
        }
        return Select-BestCatalogProduct -Products $exactList
    }

    $detectedTokens = @(Get-MatchTokens -NormalizedName $NormalizedName)
    if ($detectedTokens.Count -eq 0) {
        return $null
    }

    # 2) Single-token inventory titles are too ambiguous for fuzzy match
    #    ("Authenticator" must not become "Synology Evidence Integrity Authenticator").
    if ($detectedTokens.Count -lt 2) {
        return $null
    }

    $candidateMap = @{} # product key -> product
    foreach ($token in $detectedTokens) {
        if (-not $Index.ByToken.ContainsKey($token)) {
            continue
        }
        # Skip indexing explosion on ultra-common leftover tokens if any slip through.
        if ($Index.ByToken[$token].Count -gt 40) {
            continue
        }
        foreach ($product in $Index.ByToken[$token]) {
            $key = if ($product.ProductId) { [string]$product.ProductId } else { [string]$product.ProductName }
            if (-not $candidateMap.ContainsKey($key)) {
                $candidateMap[$key] = $product
            }
        }
    }

    if ($candidateMap.Count -eq 0) {
        return $null
    }

    $best = $null
    $bestScore = -1
    $detectedTokenSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($t in $detectedTokens) {
        [void]$detectedTokenSet.Add($t)
    }

    foreach ($product in $candidateMap.Values) {
        $catalogNormalized = [string]$product.NormalizedName
        if ([string]::IsNullOrWhiteSpace($catalogNormalized)) {
            continue
        }

        if ($NormalizedName -eq $catalogNormalized) {
            if (1000 -gt $bestScore) {
                $bestScore = 1000
                $best = $product
            }
            continue
        }

        $catalogTokens = @(Get-MatchTokens -NormalizedName $catalogNormalized)
        if ($catalogTokens.Count -eq 0) {
            continue
        }

        if (-not (Test-ProductIdentityCompatible -DetectedTokens $detectedTokens -CatalogTokens $catalogTokens)) {
            continue
        }

        $catalogTokenSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($ct in $catalogTokens) {
            [void]$catalogTokenSet.Add($ct)
        }

        # Every detected match token must appear in the catalog name (detected ⊆ catalog).
        $missingDetected = $false
        foreach ($dt in $detectedTokens) {
            if (-not $catalogTokenSet.Contains($dt)) {
                $missingDetected = $true
                break
            }
        }
        if ($missingDetected) {
            continue
        }

        $overlap = 0
        $coreOverlap = 0
        foreach ($ct in $catalogTokens) {
            if ($detectedTokenSet.Contains($ct)) {
                $overlap++
                $isIdentityOnly = (Test-IsProductYearToken -Token $ct) -or
                    (Test-IsProductMinorToken -Token $ct) -or
                    (Test-IsChannelIdentityToken -Token $ct)
                if (-not $isIdentityOnly) {
                    $coreOverlap++
                }
            }
        }

        if ($overlap -eq 0 -or $coreOverlap -eq 0) {
            continue
        }

        # Catalog must not be much broader than detected (stops "realtek"⊂"…driver…" style if tokens align poorly,
        # and "authenticator"⊂ long Synology title when single-token already blocked).
        # Require covering most catalog tokens, or catalog only slightly longer than detected.
        $ratioCatalog = $overlap / [double]$catalogTokens.Count
        $ratioDetected = $overlap / [double]$detectedTokens.Count
        $catalogExtra = $catalogTokens.Count - $detectedTokens.Count

        if ($ratioCatalog -lt 0.75 -and $catalogExtra -gt 1) {
            continue
        }
        if ($catalogTokens.Count -ge 3 -and $detectedTokens.Count -eq 2 -and $ratioCatalog -lt 0.67) {
            continue
        }

        # Prefer high coverage of both sides.
        $score = [int](($ratioCatalog * 500) + ($ratioDetected * 300) + ($coreOverlap * 80) + ($overlap * 20))

        if ($score -gt $bestScore) {
            $bestScore = $score
            $best = $product
        }
    }

    # Require a solid multi-token agreement.
    if ($bestScore -lt 500) {
        return $null
    }

    return $best
}

function New-EAMAppFromCatalogPackage {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $PackageId,

        [Parameter(Mandatory = $true)]
        [string]
        $ProductName,

        [Parameter(Mandatory = $false)]
        [string]
        $GroupId,

        [Parameter(Mandatory = $false)]
        [ValidateSet('hideAll', 'showReboot', 'showAll')]
        [string]
        $Notifications = 'hideAll',

        [Parameter(Mandatory = $false)]
        [ValidateSet('notConfigured', 'foreground')]
        [string]
        $DeliveryOptimizationPriority = 'notConfigured'
    )

    if ([string]::IsNullOrWhiteSpace($PackageId)) {
        throw "Catalog package id is empty for '$ProductName'. Cannot create app."
    }

    # Same convert + create pattern as Invoke-EAMAutoUpdate.ps1 (exclude fields that break create).
    $convertUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/convertFromMobileAppCatalogPackage(mobileAppCatalogPackageId='$PackageId')"
    Write-Host "  Converting catalog package $PackageId ..."
    $mobileAppPayload = (Invoke-MgGraphRequest -Method GET -Uri $convertUri -OutputType PSObject) |
        Select-Object * -ExcludeProperty @(
            '@odata.context'
            'id'
            'largeIcon'
            'createdDateTime'
            'lastModifiedDateTime'
            'owner'
            'notes'
            'size'
            'minimumSupportedOperatingSystem'
            'minimumFreeDiskSpaceInMB'
            'minimumMemoryInMB'
            'minimumNumberOfProcessors'
            'minimumCpuSpeedInMHz'
        )

    if ($null -eq $mobileAppPayload) {
        throw "Catalog conversion returned empty payload for package $PackageId ($ProductName)."
    }

    $body = $mobileAppPayload | ConvertTo-Json -Depth 20
    if ([string]::IsNullOrWhiteSpace($body)) {
        throw "Catalog conversion produced empty JSON for package $PackageId ($ProductName)."
    }

    Write-Host "  Creating win32CatalogApp for '$ProductName' ..."
    $created = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps' -Body $body -ContentType 'application/json' -OutputType PSObject

    $createdVersion = Get-GraphPropertyValue -Object $created -Names @('displayVersion', 'version')
    Write-Host "  Created EAM app '$($created.displayName)' Id=$($created.id) Version=$createdVersion"

    if ($GroupId) {
        $assignment = @{
            '@odata.type' = '#microsoft.graph.mobileAppAssignment'
            intent        = 'available'
            target        = @{
                '@odata.type' = '#microsoft.graph.groupAssignmentTarget'
                groupId       = $GroupId
            }
            settings      = @{
                '@odata.type'                 = '#microsoft.graph.win32CatalogAppAssignmentSettings'
                notifications                = $Notifications
                deliveryOptimizationPriority = $DeliveryOptimizationPriority
                autoUpdateSettings           = @{
                    '@odata.type'                 = '#microsoft.graph.win32LobAppAutoUpdateSettings'
                    autoUpdateSupersededAppsState = 'enabled'
                }
            }
        }

        $assignBody = $assignment | ConvertTo-Json -Depth 10 -Compress
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($created.id)/assignments" -Body $assignBody -ContentType 'application/json' | Out-Null
        Write-Host "  Assigned as available to group $GroupId (auto-update enabled, notifications=$Notifications, deliveryOptimization=$DeliveryOptimizationPriority)."
    }
    else {
        Write-Host '  Created unassigned (no group deployment).'
    }

    return $created
}

function Read-YesNo {
    param (
        [string]$Prompt,
        [bool]$DefaultYes = $false
    )

    $suffix = if ($DefaultYes) { '[Y/n]' } else { '[y/N]' }
    $answer = Read-Host -Prompt "$Prompt $suffix"
    if ([string]::IsNullOrWhiteSpace($answer)) {
        return $DefaultYes
    }
    return ($answer -match '^(y|yes)$')
}

#endregion Helpers

#region Main

Connect-EAMDiscoveryGraph

$catalogProducts = @(Get-EAMCatalogProducts)
$managedEamApps = @(Get-ManagedWin32CatalogApps)
$detectedApps = @(Get-IntuneDetectedApps)

$excludeSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($name in $ExcludeCatalogNames) {
    if ($name) { [void]$excludeSet.Add($name.Trim()) }
}

Write-Host 'Building catalog match index...'
$catalogIndex = New-CatalogMatchIndex -CatalogProducts $catalogProducts
Write-Host "  Exact keys: $($catalogIndex.Exact.Count); token keys: $($catalogIndex.ByToken.Count)"

Write-Host 'Mapping existing Intune EAM apps to catalog products...'
$managedCatalogLookup = Build-ManagedCatalogLookup `
    -ManagedApps $managedEamApps `
    -CatalogProducts $catalogProducts `
    -CatalogIndex $catalogIndex

# Aggregate detected inventory first (unique normalized titles), then match once per title.
Write-Host 'Aggregating detected apps by normalized name...'
$detectedGroups = @{} # normalizedName -> aggregate
foreach ($detected in $detectedApps) {
    $displayName = [string]$detected.displayName
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        continue
    }

    $deviceCount = 0
    if ($null -ne $detected.deviceCount) {
        $deviceCount = [int]$detected.deviceCount
    }
    if ($deviceCount -le 0) {
        continue
    }

    $normalizedStrict = Get-NormalizedAppName -Name $displayName
    $normalizedSoft = Get-NormalizedAppName -Name $displayName -KeepStopwords
    $normalized = if ($normalizedStrict) { $normalizedStrict } else { $normalizedSoft }
    if (-not $normalized) {
        continue
    }

    if (-not $detectedGroups.ContainsKey($normalized)) {
        $detectedGroups[$normalized] = [PSCustomObject]@{
            NormalizedName       = $normalized
            TotalDevices         = 0
            Versions             = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            DetectedDisplayNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            VersionDeviceMap     = @{}
        }
    }

    $group = $detectedGroups[$normalized]
    $group.TotalDevices += $deviceCount
    $version = if ($detected.version) { [string]$detected.version } else { 'unknown' }
    [void]$group.Versions.Add($version)
    [void]$group.DetectedDisplayNames.Add($displayName)
    if (-not $group.VersionDeviceMap.ContainsKey($version)) {
        $group.VersionDeviceMap[$version] = 0
    }
    $group.VersionDeviceMap[$version] += $deviceCount
}

Write-Host "  Unique detected titles: $($detectedGroups.Count) (from $($detectedApps.Count) inventory rows)"

Write-Host 'Matching unique titles to EAM catalog products...'
$matchBuckets = @{} # catalog product key -> aggregated match
$matchCache = @{}   # normalized detected name -> catalog product or $false
$processed = 0
$totalTitles = $detectedGroups.Count
$sw = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($groupEntry in $detectedGroups.GetEnumerator()) {
    $processed++
    if ($processed % 500 -eq 0 -or $processed -eq $totalTitles) {
        Write-Host ("  ... {0}/{1} titles ({2:n1}s)" -f $processed, $totalTitles, $sw.Elapsed.TotalSeconds)
    }

    $group = $groupEntry.Value
    # Skip sparse titles early (they cannot meet MinDeviceCount after aggregation either).
    if ($group.TotalDevices -lt $MinDeviceCount) {
        continue
    }

    $normalized = $group.NormalizedName
    if ($matchCache.ContainsKey($normalized)) {
        $match = $matchCache[$normalized]
        if ($match -eq $false) {
            continue
        }
    }
    else {
        $match = Find-CatalogMatchFromIndex -NormalizedName $normalized -Index $catalogIndex
        $matchCache[$normalized] = if ($match) { $match } else { $false }
        if (-not $match) {
            continue
        }
    }

    if ($excludeSet.Contains($match.ProductName)) {
        continue
    }

    $key = if ($match.ProductId) { [string]$match.ProductId } else { [string]$match.ProductName }
    if (-not $matchBuckets.ContainsKey($key)) {
        $matchBuckets[$key] = [PSCustomObject]@{
            Catalog        = $match
            # One contribution per detected title (normalized). Do NOT blindly sum all titles —
            # that double-counts / inflates when fuzzy match merges unrelated software.
            Contributions  = [System.Collections.Generic.List[object]]::new()
        }
    }

    $bucket = $matchBuckets[$key]
    $versionMapCopy = @{}
    foreach ($vk in $group.VersionDeviceMap.Keys) {
        $versionMapCopy[$vk] = [int]$group.VersionDeviceMap[$vk]
    }
    $displayNamesCopy = [System.Collections.Generic.List[string]]::new()
    foreach ($dn in $group.DetectedDisplayNames) {
        $displayNamesCopy.Add([string]$dn)
    }

    $isExactNameMatch = (
        $group.NormalizedName -and $match.NormalizedName -and
        $group.NormalizedName -eq $match.NormalizedName
    )

    $bucket.Contributions.Add([PSCustomObject]@{
            NormalizedName       = $group.NormalizedName
            TotalDevices         = [int]$group.TotalDevices
            Versions             = @($group.Versions)
            VersionDeviceMap     = $versionMapCopy
            DetectedDisplayNames = $displayNamesCopy
            ExactNameMatch       = $isExactNameMatch
        })
}

Write-Host ("  Catalog matches: {0} product(s) in {1:n1}s" -f $matchBuckets.Count, $sw.Elapsed.TotalSeconds)

Write-Host 'Scoring candidates (version fragmentation / catalog lag)...'
$scoreSw = [System.Diagnostics.Stopwatch]::StartNew()
$candidates = [System.Collections.Generic.List[object]]::new()
$scoreIndex = 0
$scoreTotal = $matchBuckets.Count

foreach ($entry in $matchBuckets.GetEnumerator()) {
    $scoreIndex++
    if ($scoreIndex % 25 -eq 0 -or $scoreIndex -eq $scoreTotal) {
        Write-Host ("  ... scored {0}/{1} ({2:n1}s)" -f $scoreIndex, $scoreTotal, $scoreSw.Elapsed.TotalSeconds)
    }

    $bucket = $entry.Value
    $catalog = $bucket.Catalog
    $contributions = @($bucket.Contributions)
    if ($contributions.Count -eq 0) {
        continue
    }

    $managedInfo = Test-CatalogProductAlreadyManaged -CatalogProduct $catalog -ManagedCatalogLookup $managedCatalogLookup
    $alreadyManaged = $null -ne $managedInfo

    if ($alreadyManaged -and -not $IncludeAlreadyManaged) {
        continue
    }

    # Primary detected title for device count / versions:
    # prefer exact name match, then highest device count for that single title.
    # Summing all fuzzy-matched titles inflates counts (same fleet counted under many names).
    $primary = $contributions |
        Sort-Object -Property `
        @{ Expression = { if ($_.ExactNameMatch) { 1 } else { 0 } }; Descending = $true }, `
        @{ Expression = 'TotalDevices'; Descending = $true } |
        Select-Object -First 1

    $deviceCount = [int]$primary.TotalDevices
    $deviceCountSumAllTitles = 0
    foreach ($c in $contributions) {
        $deviceCountSumAllTitles += [int]$c.TotalDevices
    }
    $matchedTitleCount = $contributions.Count

    if ($deviceCount -lt $MinDeviceCount) {
        continue
    }

    $versionList = [System.Collections.Generic.List[string]]::new()
    foreach ($v in @($primary.Versions)) {
        $versionList.Add([string]$v)
    }
    $versionList.Sort([System.StringComparer]::OrdinalIgnoreCase)

    $distinctVersions = $versionList.Count
    if ($distinctVersions -lt $MinDistinctVersions) {
        continue
    }

    # Cap summary length so hosts do not choke on huge multi-version strings
    $summaryParts = [System.Collections.Generic.List[string]]::new()
    $maxVersionParts = 12
    for ($vi = 0; $vi -lt $versionList.Count -and $vi -lt $maxVersionParts; $vi++) {
        $v = $versionList[$vi]
        $count = 0
        if ($primary.VersionDeviceMap.ContainsKey($v)) {
            $count = [int]$primary.VersionDeviceMap[$v]
        }
        $summaryParts.Add("$v ($count)")
    }
    if ($versionList.Count -gt $maxVersionParts) {
        $summaryParts.Add("+$( $versionList.Count - $maxVersionParts ) more")
    }
    $versionSummary = $summaryParts -join '; '

    # Version lag: any detected version strictly older than catalog latest?
    $lagCount = 0
    $newestDetected = $null
    foreach ($v in $versionList) {
        if ($v -eq 'unknown') {
            continue
        }
        if (-not $newestDetected) {
            $newestDetected = $v
        }
        else {
            $cmpNew = Compare-VersionString -Left $v -Right $newestDetected
            if ($cmpNew -eq 1) {
                $newestDetected = $v
            }
        }

        $cmpLag = Compare-VersionString -Left $v -Right ([string]$catalog.LatestVersion)
        if ($cmpLag -eq -1) {
            $lagCount++
        }
    }

    $behindCatalog = $false
    if ($newestDetected -and $catalog.LatestVersion) {
        $cmpBehind = Compare-VersionString -Left $newestDetected -Right ([string]$catalog.LatestVersion)
        $behindCatalog = ($cmpBehind -eq -1)
    }

    # Priority uses primary-title device count (not inflated multi-title sum).
    $priorityScore =
        $deviceCount +
        ($distinctVersions * 40) +
        ($(if ($behindCatalog) { 80 } else { 0 })) +
        ($lagCount * 15)

    $selfUpdateRisk = if ($distinctVersions -ge 3 -or ($distinctVersions -ge 2 -and $behindCatalog)) {
        'High'
    }
    elseif ($distinctVersions -ge 2 -or $behindCatalog) {
        'Medium'
    }
    else {
        'Low'
    }

    $nameSamples = [System.Collections.Generic.List[string]]::new()
    foreach ($dn in $primary.DetectedDisplayNames) {
        $nameSamples.Add([string]$dn)
        if ($nameSamples.Count -ge 5) {
            break
        }
    }
    # Note other titles that also matched this catalog product (possible over-match).
    $otherTitles = [System.Collections.Generic.List[string]]::new()
    foreach ($c in $contributions) {
        if ($c -eq $primary) {
            continue
        }
        foreach ($dn in $c.DetectedDisplayNames) {
            if ($otherTitles.Count -ge 5) {
                break
            }
            if (-not $nameSamples.Contains([string]$dn)) {
                $otherTitles.Add([string]$dn)
            }
        }
    }

    $recommendation = switch ($selfUpdateRisk) {
        'High' { 'Prioritize for EAM: multiple versions / lag suggests installs do not self-update consistently.' }
        'Medium' { 'Consider EAM: some version spread or catalog lag.' }
        default { 'Optional: present on many devices; single version may already be controlled or auto-updating.' }
    }
    if ($matchedTitleCount -gt 3 -or ($deviceCountSumAllTitles -gt ($deviceCount * 2) -and $matchedTitleCount -gt 1)) {
        $recommendation += " Warning: $matchedTitleCount detected titles mapped here (sum would be $deviceCountSumAllTitles); DeviceCount uses the primary title only to avoid inflation."
    }

    $candidates.Add([PSCustomObject]@{
            PriorityScore            = $priorityScore
            SelfUpdateRisk           = $selfUpdateRisk
            CatalogProductName       = [string]$catalog.ProductName
            CatalogPublisher         = [string]$catalog.Publisher
            CatalogLatestVersion     = [string]$catalog.LatestVersion
            CatalogPackageId         = [string]$catalog.LatestPackageId
            CatalogProductId         = [string]$catalog.ProductId
            DeviceCount              = $deviceCount
            DeviceCountSumAllTitles  = $deviceCountSumAllTitles
            MatchedDetectedTitleCount = $matchedTitleCount
            PrimaryDetectedTitle     = if ($primary.DetectedDisplayNames.Count -gt 0) { [string]$primary.DetectedDisplayNames[0] } else { $primary.NormalizedName }
            DistinctVersions         = $distinctVersions
            DetectedVersions         = $versionSummary
            NewestDetectedVersion    = $newestDetected
            BehindCatalogLatest      = $behindCatalog
            AlreadyManagedAsEAM      = [bool]$alreadyManaged
            ManagedIntuneAppName     = if ($managedInfo) { [string]$managedInfo.ManagedAppName } else { $null }
            ManagedIntuneAppId       = if ($managedInfo) { [string]$managedInfo.ManagedAppId } else { $null }
            DetectedNameSamples      = ($nameSamples -join ' | ')
            OtherMatchedTitles       = ($otherTitles -join ' | ')
            Recommendation           = $recommendation
        })
}

Write-Host ("  Candidates after filters: {0} ({1:n1}s)" -f $candidates.Count, $scoreSw.Elapsed.TotalSeconds)

$sorted = @($candidates | Sort-Object -Property @{ Expression = 'PriorityScore'; Descending = $true }, @{ Expression = 'DeviceCount'; Descending = $true })
if ($Top -gt 0 -and $sorted.Count -gt $Top) {
    $sorted = @($sorted | Select-Object -First $Top)
}

Write-Host ''
Write-Host "Potential EAM candidates: $($sorted.Count) (MinDeviceCount=$MinDeviceCount, MinDistinctVersions=$MinDistinctVersions)"
Write-Host ''

if ($sorted.Count -eq 0) {
    Write-Host 'No candidates matched the catalog with the current filters.'
    return @()
}

# Avoid Format-Table -AutoSize (can appear to hang in some hosts on wide objects)
$sorted | Format-Table PriorityScore, SelfUpdateRisk, CatalogProductName, DeviceCount, MatchedDetectedTitleCount, DeviceCountSumAllTitles, DistinctVersions, CatalogLatestVersion, NewestDetectedVersion, AlreadyManagedAsEAM, PrimaryDetectedTitle -Wrap | Out-String -Width 260 | Write-Host

Write-Host 'DeviceCount = installs for the primary detected title only (not a unique-device census).'
Write-Host 'DeviceCountSumAllTitles = sum across every detected title that mapped to the catalog product (often inflated).'
Write-Host ''
Write-Host 'Version detail (top results):'
foreach ($row in ($sorted | Select-Object -First ([Math]::Min(15, $sorted.Count)))) {
    Write-Host ("- {0}: devices={1} (sumAllTitles={2}, matchedTitles={3}); primary='{4}'; versions={5}" -f `
            $row.CatalogProductName, $row.DeviceCount, $row.DeviceCountSumAllTitles, $row.MatchedDetectedTitleCount, `
            $row.PrimaryDetectedTitle, $row.DetectedVersions)
    if ($row.OtherMatchedTitles) {
        Write-Host ("    other mapped titles: {0}" -f $row.OtherMatchedTitles)
    }
}

if ($ExportPath) {
    $sorted | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host ''
    Write-Host "Exported report to $ExportPath"
}

if (-not $AddApps) {
    Write-Host ''
    Write-Host 'Discovery only. Re-run with -AddApps to create EAM apps from selected candidates.'
    return $sorted
}

Write-Host ''
Write-Host '--- Add EAM apps ---'
Write-Host 'You will be prompted Y/n for each candidate (unless -AddAll or -AddAppNames was used).'
try { [Console]::Out.Flush() } catch { }

$toAdd = [System.Collections.Generic.List[object]]::new()
$addNameSet = $null
if ($AddAppNames -and $AddAppNames.Count -gt 0) {
    $addNameSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($n in $AddAppNames) {
        if ($n) { [void]$addNameSet.Add($n.Trim()) }
    }
}

$promptIndex = 0
foreach ($candidate in $sorted) {
    if ($candidate.AlreadyManagedAsEAM -and -not $IncludeAlreadyManaged) {
        continue
    }

    if ($addNameSet) {
        if (-not $addNameSet.Contains($candidate.CatalogProductName)) {
            continue
        }
        $toAdd.Add($candidate)
        continue
    }

    if ($AddAll) {
        $toAdd.Add($candidate)
        continue
    }

    $promptIndex++
    $prompt = "[$promptIndex/$($sorted.Count)] Add EAM app '$($candidate.CatalogProductName)' (devices=$($candidate.DeviceCount), risk=$($candidate.SelfUpdateRisk), catalog=$($candidate.CatalogLatestVersion))?"
    Write-Host $prompt
    try { [Console]::Out.Flush() } catch { }
    if (Read-YesNo -Prompt 'Add this app' -DefaultYes $false) {
        $toAdd.Add($candidate)
    }
}

if ($toAdd.Count -eq 0) {
    Write-Host 'No apps selected for add.'
    return $sorted
}

Write-Host "Adding $($toAdd.Count) app(s)..."
$addResults = [System.Collections.Generic.List[object]]::new()

foreach ($candidate in $toAdd) {
    if (-not $candidate.CatalogPackageId) {
        Write-Warning "Skipping '$($candidate.CatalogProductName)': no catalog package id."
        continue
    }

    $targetLabel = $candidate.CatalogProductName
    if ($WhatIfAdd -or $WhatIfPreference) {
        Write-Host "WhatIf: would create EAM app for '$targetLabel' from package $($candidate.CatalogPackageId)"
        $addResults.Add([PSCustomObject]@{
                CatalogProductName = $targetLabel
                Action             = 'WhatIf'
                AppId              = $null
                Succeeded          = $true
                Error              = $null
            })
        continue
    }

    if ($PSCmdlet.ShouldProcess($targetLabel, 'Create win32CatalogApp from EAM catalog package')) {
        try {
            $created = New-EAMAppFromCatalogPackage `
                -PackageId $candidate.CatalogPackageId `
                -ProductName $candidate.CatalogProductName `
                -GroupId $AssignGroupId `
                -Notifications $AssignmentNotifications `
                -DeliveryOptimizationPriority $AssignmentDeliveryOptimization

            $addResults.Add([PSCustomObject]@{
                    CatalogProductName = $targetLabel
                    Action             = 'Created'
                    AppId              = $created.id
                    DisplayVersion     = $created.displayVersion
                    Succeeded          = $true
                    Error              = $null
                })
        }
        catch {
            Write-Error "Failed to add '$targetLabel': $_"
            $addResults.Add([PSCustomObject]@{
                    CatalogProductName = $targetLabel
                    Action             = 'Failed'
                    AppId              = $null
                    Succeeded          = $false
                    Error              = $_.Exception.Message
                })
        }
    }
}

Write-Host ''
Write-Host 'Add summary:'
$addResults | Format-Table -AutoSize | Out-Host

return [PSCustomObject]@{
    Candidates = $sorted
    AddResults = $addResults
}

#endregion Main
