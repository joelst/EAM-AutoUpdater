#Requires -Version 7.0
#Requires -Modules Az.Accounts

<#
.SYNOPSIS
    Adds required modules or updates existing packages on an Azure Automation runtime environment.

.DESCRIPTION
    Two modes:

    Ensure required (default)
      Imports the Microsoft Graph modules required by EAM-AutoUpdater. Missing packages
      are added; existing ones are re-imported unless -SkipExisting is set.

    Update existing (-UpdateExistingModules)
      Re-imports packages already on the runtime from the PowerShell Gallery.
      - Omit -UpdateModuleNames to target every non-default installed package.
      - Pass -UpdateModuleNames to update only the listed packages (must already exist).
      Prompts Y/n for each module unless -Force is specified.

    Uses the Automation packages REST API (api-version 2024-10-23). Gallery import is
    asynchronous; use -Wait to poll until each package reaches Succeeded or Failed.

    Before updating an already-installed package, the script queries the PowerShell Gallery
    and skips the import (and any prompt) when the installed version is already current
    relative to the target gallery version. Use -SkipVersionCheck to force re-import.

.PARAMETER ResourceGroupName
    Resource group that contains the Automation Account.

.PARAMETER AutomationAccountName
    Automation Account name.

.PARAMETER RuntimeEnvironmentName
    Target runtime environment name (e.g. EAM-PS72-Graph).

.PARAMETER ModuleNames
    Modules to ensure (add/update) in the default mode. Defaults to the EAM-AutoUpdater required set.
    Authentication is always processed first when present in the list.

.PARAMETER UpdateExistingModules
    Switch to the update-existing path. Does not add missing required modules; only
    re-imports packages already present on the runtime.

.PARAMETER UpdateModuleNames
    With -UpdateExistingModules: optional list of package names to update.
    If omitted, all non-default packages currently on the runtime are candidates.

.PARAMETER ModuleVersions
    Optional hashtable of module name to gallery version (e.g. @{ 'Microsoft.Graph.Authentication' = '2.25.0' }).
    Omit a module (or leave empty) to use the latest gallery package URI.

.PARAMETER SubscriptionId
    Optional subscription ID. Defaults to the current Az context subscription.

.PARAMETER SkipExisting
    Ensure-required mode only. When set, modules already present (except Failed) are left unchanged.

.PARAMETER Force
    Update-existing mode: do not prompt for each module; update all selected packages
    that have a newer gallery version (version precheck still applies unless -SkipVersionCheck).

.PARAMETER SkipVersionCheck
    Always re-import selected packages even when the installed version matches the gallery
    (or pinned) target version. Useful for repairing a broken package.

.PARAMETER Wait
    Wait for each import to finish (Succeeded or Failed) before continuing.
    Recommended so Authentication is fully available before dependent modules import.

.PARAMETER WaitTimeoutMinutes
    Max time to wait per module when -Wait is specified. Default: 30.

.PARAMETER PollIntervalSeconds
    Seconds between provisioning state polls when -Wait is specified. Default: 15.

.EXAMPLE
    # Ensure required Graph modules (add missing / update existing on the list)
    .\Update-RuntimeEnvironmentModules.ps1 `
        -ResourceGroupName 'rg-eam-autoupdater' `
        -AutomationAccountName 'aa-eam-autoupdater' `
        -RuntimeEnvironmentName 'EAM-PS72-Graph' `
        -Wait

.EXAMPLE
    # Update ALL non-default packages already on the runtime (prompt per module)
    .\Update-RuntimeEnvironmentModules.ps1 `
        -ResourceGroupName 'rg-eam' `
        -AutomationAccountName 'aa-eam' `
        -RuntimeEnvironmentName 'EAM-PS72-Graph' `
        -UpdateExistingModules `
        -Wait

.EXAMPLE
    # Update only a specified list of existing packages (prompt per module)
    .\Update-RuntimeEnvironmentModules.ps1 `
        -ResourceGroupName 'rg-eam' `
        -AutomationAccountName 'aa-eam' `
        -RuntimeEnvironmentName 'EAM-PS72-Graph' `
        -UpdateExistingModules `
        -UpdateModuleNames 'Microsoft.Graph.Authentication','Microsoft.Graph.Groups' `
        -Wait

.EXAMPLE
    # Update all existing packages without prompts (still skips packages already current)
    .\Update-RuntimeEnvironmentModules.ps1 `
        -ResourceGroupName 'rg-eam' `
        -AutomationAccountName 'aa-eam' `
        -RuntimeEnvironmentName 'EAM-PS72-Graph' `
        -UpdateExistingModules `
        -Force `
        -Wait

.EXAMPLE
    # Force re-import even when gallery reports the same version
    .\Update-RuntimeEnvironmentModules.ps1 `
        -ResourceGroupName 'rg-eam' `
        -AutomationAccountName 'aa-eam' `
        -RuntimeEnvironmentName 'EAM-PS72-Graph' `
        -UpdateExistingModules `
        -UpdateModuleNames 'Microsoft.Graph.Authentication' `
        -SkipVersionCheck `
        -Force `
        -Wait
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'EnsureRequired')]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $ResourceGroupName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $AutomationAccountName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $RuntimeEnvironmentName,

    [Parameter(Mandatory = $false, ParameterSetName = 'EnsureRequired')]
    [string[]]
    $ModuleNames = @(
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Beta.DeviceManagement.Actions'
        'Microsoft.Graph.Beta.Devices.CorporateManagement'
        'Microsoft.Graph.Groups'
        'Microsoft.Graph.Beta.DeviceManagement'
    ),

    [Parameter(Mandatory = $true, ParameterSetName = 'UpdateExisting')]
    [switch]
    $UpdateExistingModules,

    [Parameter(Mandatory = $false, ParameterSetName = 'UpdateExisting')]
    [string[]]
    $UpdateModuleNames,

    [Parameter(Mandatory = $false)]
    [hashtable]
    $ModuleVersions = @{},

    [Parameter(Mandatory = $false)]
    [string]
    $SubscriptionId,

    [Parameter(Mandatory = $false, ParameterSetName = 'EnsureRequired')]
    [switch]
    $SkipExisting,

    [Parameter(Mandatory = $false)]
    [switch]
    $Force,

    [Parameter(Mandatory = $false)]
    [switch]
    $SkipVersionCheck,

    [Parameter(Mandatory = $false)]
    [switch]
    $Wait,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 180)]
    [int]
    $WaitTimeoutMinutes = 30,

    [Parameter(Mandatory = $false)]
    [ValidateRange(5, 120)]
    [int]
    $PollIntervalSeconds = 15
)

$ErrorActionPreference = 'Stop'
$apiVersion = '2024-10-23'
$authName = 'Microsoft.Graph.Authentication'

function Test-AzContextValid {
    param (
        [Parameter(Mandatory = $false)]
        [object]
        $Context
    )

    if (-not $Context) {
        return $false
    }

    if (-not $Context.Account -or [string]::IsNullOrWhiteSpace([string]$Context.Account.Id)) {
        return $false
    }

    if (-not $Context.Subscription -or [string]::IsNullOrWhiteSpace([string]$Context.Subscription.Id)) {
        return $false
    }

    return $true
}

function Get-RequiredAzContext {
    $context = Get-AzContext -ErrorAction SilentlyContinue

    if (Test-AzContextValid -Context $context) {
        Write-Host "Using Azure context: $($context.Account.Id) / subscription '$($context.Subscription.Name)' ($($context.Subscription.Id))"
        return $context
    }

    if ($context -and $context.Account -and -not $context.Subscription) {
        Write-Warning 'An Azure account is signed in but no subscription is selected.'
    }
    else {
        Write-Warning 'No valid Azure context found (Connect-AzAccount has not been run, or the session expired).'
    }

    $prompt = if ($context -and $context.Account -and -not $context.Subscription) {
        'Run Connect-AzAccount / select a subscription now? [Y/n]'
    }
    else {
        'Run Connect-AzAccount now? [Y/n]'
    }

    $answer = Read-Host -Prompt $prompt
    if ([string]::IsNullOrWhiteSpace($answer)) {
        $answer = 'Y'
    }

    if ($answer -notmatch '^(y|yes)$') {
        throw 'Azure authentication required. Run Connect-AzAccount, select a subscription, then re-run this script.'
    }

    Write-Host 'Starting Connect-AzAccount...'
    $null = Connect-AzAccount -ErrorAction Stop

    $context = Get-AzContext -ErrorAction SilentlyContinue
    if (-not (Test-AzContextValid -Context $context)) {
        $subscriptions = @(Get-AzSubscription -ErrorAction SilentlyContinue)
        if ($subscriptions.Count -eq 0) {
            throw 'Connect-AzAccount completed but no subscriptions are available for this identity.'
        }

        if ($subscriptions.Count -eq 1) {
            $null = Set-AzContext -SubscriptionId $subscriptions[0].Id -ErrorAction Stop
        }
        else {
            Write-Host 'Select a subscription:'
            for ($i = 0; $i -lt $subscriptions.Count; $i++) {
                Write-Host ("  [{0}] {1} ({2})" -f $i, $subscriptions[$i].Name, $subscriptions[$i].Id)
            }

            $selection = Read-Host -Prompt "Enter subscription number (0-$($subscriptions.Count - 1))"
            $index = 0
            if (-not [int]::TryParse($selection, [ref]$index) -or $index -lt 0 -or $index -ge $subscriptions.Count) {
                throw "Invalid subscription selection '$selection'."
            }

            $null = Set-AzContext -SubscriptionId $subscriptions[$index].Id -ErrorAction Stop
        }

        $context = Get-AzContext -ErrorAction SilentlyContinue
    }

    if (-not (Test-AzContextValid -Context $context)) {
        throw 'Azure context is still invalid after Connect-AzAccount. Verify your sign-in and subscription access.'
    }

    Write-Host "Connected as $($context.Account.Id) / subscription '$($context.Subscription.Name)' ($($context.Subscription.Id))"
    return $context
}

function Get-GalleryContentUri {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ModuleName,

        [Parameter(Mandatory = $false)]
        [string]
        $Version
    )

    if ([string]::IsNullOrWhiteSpace($Version)) {
        return "https://www.powershellgallery.com/api/v2/package/$ModuleName"
    }

    return "https://www.powershellgallery.com/api/v2/package/$ModuleName/$Version"
}

function ConvertTo-ComparableVersion {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $VersionString
    )

    $normalized = $VersionString.Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    # Prefer semantic version (supports prerelease labels used by some modules).
    try {
        return [System.Management.Automation.SemanticVersion]::Parse($normalized)
    }
    catch {
        # Fall back to System.Version after stripping a common prerelease suffix.
    }

    $core = ($normalized -split '-', 2)[0]
    try {
        return [version]$core
    }
    catch {
        return $null
    }
}

function Compare-PackageVersionString {
    <#
    .SYNOPSIS
        Compares two module version strings.
    .OUTPUTS
        int: 1 if Left > Right, 0 if equal, -1 if Left < Right, $null if incomparable.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]
        $Left,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]
        $Right
    )

    if ([string]::IsNullOrWhiteSpace($Left) -or [string]::IsNullOrWhiteSpace($Right)) {
        return $null
    }

    if ($Left.Trim() -eq $Right.Trim()) {
        return 0
    }

    $leftVersion = ConvertTo-ComparableVersion -VersionString $Left
    $rightVersion = ConvertTo-ComparableVersion -VersionString $Right
    if ($null -eq $leftVersion -or $null -eq $rightVersion) {
        return $null
    }

    if ($leftVersion -gt $rightVersion) {
        return 1
    }
    if ($leftVersion -lt $rightVersion) {
        return -1
    }

    return 0
}

function Get-XmlChildInnerText {
    param (
        [Parameter(Mandatory = $false)]
        [System.Xml.XmlNode]
        $Parent,

        [Parameter(Mandatory = $true)]
        [string]
        $LocalName
    )

    if (-not $Parent) {
        return $null
    }

    $child = $Parent.ChildNodes |
        Where-Object { $_.LocalName -eq $LocalName } |
        Select-Object -First 1

    if (-not $child) {
        return $null
    }

    return [string]$child.InnerText
}

function Get-PSGalleryEntryVersion {
    param (
        [Parameter(Mandatory = $true)]
        $Entry
    )

    # Prefer OData properties (namespaced d:Version / d:NormalizedVersion).
    if ($Entry -is [System.Xml.XmlNode] -or $Entry.properties) {
        $props = $Entry.properties
        $fromProps = Get-XmlChildInnerText -Parent $props -LocalName 'NormalizedVersion'
        if ([string]::IsNullOrWhiteSpace($fromProps)) {
            $fromProps = Get-XmlChildInnerText -Parent $props -LocalName 'Version'
        }
        if (-not [string]::IsNullOrWhiteSpace($fromProps)) {
            return $fromProps.Trim()
        }
    }

    # Fallback: Packages(Id='Name',Version='x.y.z') in the entry id.
    $entryId = [string]$Entry.id
    if ($entryId -match "Version='([^']+)'") {
        return $Matches[1].Trim()
    }

    # Fallback: content src URL .../package/Name/x.y.z
    $contentSrc = $null
    if ($Entry.content -and $Entry.content.src) {
        $contentSrc = [string]$Entry.content.src
    }
    elseif ($Entry.content -is [System.Xml.XmlElement] -and $Entry.content.GetAttribute('src')) {
        $contentSrc = $Entry.content.GetAttribute('src')
    }

    if ($contentSrc -match '/([^/]+)$') {
        return $Matches[1].Trim()
    }

    return $null
}

function Get-PSGalleryModuleVersion {
    <#
    .SYNOPSIS
        Resolves a module version from the PowerShell Gallery (pinned or latest stable).
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ModuleName,

        [Parameter(Mandatory = $false)]
        [string]
        $PinnedVersion
    )

    if (-not [string]::IsNullOrWhiteSpace($PinnedVersion)) {
        return [PSCustomObject]@{
            ModuleName = $ModuleName
            Version    = $PinnedVersion.Trim()
            Source     = 'Pinned'
            Resolved   = $true
            Error      = $null
        }
    }

    $escapedId = $ModuleName.Replace("'", "''")
    # IsLatestVersion excludes prerelease packages.
    $filter = "Id eq '$escapedId' and IsLatestVersion"
    $uri = "https://www.powershellgallery.com/api/v2/Packages()?`$filter=$([uri]::EscapeDataString($filter))"

    try {
        $response = Invoke-RestMethod -Uri $uri -Method Get -TimeoutSec 60 -ErrorAction Stop

        # Single match: root is often the Atom <entry>. Multiple matches: feed.entry.
        $entry = $null
        if ($null -ne $response.feed) {
            $entry = $response.feed.entry
        }
        elseif ($null -ne $response.entry) {
            $entry = $response.entry
        }
        elseif ($response -is [System.Xml.XmlElement] -and $response.LocalName -eq 'entry') {
            $entry = $response
        }
        else {
            $entry = $response
        }

        if ($entry -is [System.Array]) {
            $entry = $entry | Select-Object -First 1
        }

        if ($null -eq $entry) {
            return [PSCustomObject]@{
                ModuleName = $ModuleName
                Version    = $null
                Source     = 'Gallery'
                Resolved   = $false
                Error      = 'No package entry returned from the PowerShell Gallery.'
            }
        }

        $version = Get-PSGalleryEntryVersion -Entry $entry
        if ([string]::IsNullOrWhiteSpace([string]$version)) {
            return [PSCustomObject]@{
                ModuleName = $ModuleName
                Version    = $null
                Source     = 'Gallery'
                Resolved   = $false
                Error      = 'Gallery response did not include a Version property.'
            }
        }

        return [PSCustomObject]@{
            ModuleName = $ModuleName
            Version    = ([string]$version).Trim()
            Source     = 'Gallery'
            Resolved   = $true
            Error      = $null
        }
    }
    catch {
        return [PSCustomObject]@{
            ModuleName = $ModuleName
            Version    = $null
            Source     = 'Gallery'
            Resolved   = $false
            Error      = $_.Exception.Message
        }
    }
}

function Invoke-AutomationRest {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'PUT', 'DELETE')]
        [string]
        $Method,

        [Parameter(Mandatory = $true)]
        [string]
        $RelativePath,

        [Parameter(Mandatory = $false)]
        [object]
        $Payload
    )

    $params = @{
        Method = $Method
        Path   = "$RelativePath`?api-version=$apiVersion"
    }

    if ($PSBoundParameters.ContainsKey('Payload') -and $null -ne $Payload) {
        $params['Payload'] = ($Payload | ConvertTo-Json -Depth 10 -Compress)
    }

    $response = Invoke-AzRestMethod @params
    if ($response.StatusCode -ge 400) {
        $bodyText = $response.Content
        throw "Azure REST $Method $RelativePath failed with status $($response.StatusCode). Content: $bodyText"
    }

    if ([string]::IsNullOrWhiteSpace($response.Content)) {
        return $null
    }

    return ($response.Content | ConvertFrom-Json)
}

function Get-RuntimeEnvironmentPackages {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $BasePath
    )

    $packages = [System.Collections.Generic.List[object]]::new()
    $path = "$BasePath/packages"
    $nextLink = $null

    do {
        if ($nextLink) {
            $response = Invoke-AzRestMethod -Method GET -Uri $nextLink
            if ($response.StatusCode -ge 400) {
                throw "Failed to list packages (next page). Status $($response.StatusCode): $($response.Content)"
            }
            $page = $response.Content | ConvertFrom-Json
        }
        else {
            $page = Invoke-AutomationRest -Method GET -RelativePath $path
        }

        foreach ($item in @($page.value)) {
            $packages.Add($item)
        }

        $nextLink = $page.nextLink
    } while ($nextLink)

    return $packages
}

function Get-PackageState {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $BasePath,

        [Parameter(Mandatory = $true)]
        [string]
        $PackageName
    )

    $encodedName = [uri]::EscapeDataString($PackageName)
    return Invoke-AutomationRest -Method GET -RelativePath "$BasePath/packages/$encodedName"
}

function Wait-PackageProvisioning {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $BasePath,

        [Parameter(Mandatory = $true)]
        [string]
        $PackageName,

        [Parameter(Mandatory = $true)]
        [int]
        $TimeoutMinutes,

        [Parameter(Mandatory = $true)]
        [int]
        $PollSeconds
    )

    $deadline = [DateTime]::UtcNow.AddMinutes($TimeoutMinutes)
    $terminalSuccess = @('Succeeded')
    $terminalFailure = @('Failed', 'Canceled')

    do {
        $package = Get-PackageState -BasePath $BasePath -PackageName $PackageName
        $state = [string]$package.properties.provisioningState
        $version = [string]$package.properties.version
        $errorMessage = [string]$package.properties.error.message

        Write-Host "  [$PackageName] provisioningState=$state$(if ($version) { " version=$version" })"

        if ($terminalSuccess -contains $state) {
            return [PSCustomObject]@{
                Name              = $PackageName
                ProvisioningState = $state
                Version           = $version
                Succeeded         = $true
                ErrorMessage      = $null
            }
        }

        if ($terminalFailure -contains $state) {
            return [PSCustomObject]@{
                Name              = $PackageName
                ProvisioningState = $state
                Version           = $version
                Succeeded         = $false
                ErrorMessage      = $errorMessage
            }
        }

        if ([DateTime]::UtcNow -ge $deadline) {
            return [PSCustomObject]@{
                Name              = $PackageName
                ProvisioningState = $state
                Version           = $version
                Succeeded         = $false
                ErrorMessage      = "Timed out after $TimeoutMinutes minute(s) waiting for package import."
            }
        }

        Start-Sleep -Seconds $PollSeconds
    } while ($true)
}

function Set-RuntimeEnvironmentPackage {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $BasePath,

        [Parameter(Mandatory = $true)]
        [string]
        $PackageName,

        [Parameter(Mandatory = $true)]
        [string]
        $ContentUri
    )

    $encodedName = [uri]::EscapeDataString($PackageName)
    $payload = @{
        properties = @{
            contentLink = @{
                uri = $ContentUri
            }
        }
    }

    return Invoke-AutomationRest -Method PUT -RelativePath "$BasePath/packages/$encodedName" -Payload $payload
}

function Get-OrderedModuleNames {
    param (
        [Parameter(Mandatory = $true)]
        [string[]]
        $Names
    )

    $ordered = [System.Collections.Generic.List[string]]::new()
    if ($Names -contains $authName) {
        $ordered.Add($authName)
    }

    foreach ($name in $Names) {
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }
        if ($name -ne $authName -and -not $ordered.Contains($name)) {
            $ordered.Add($name)
        }
    }

    return $ordered
}

function Read-YesNo {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $Prompt,

        [Parameter(Mandatory = $false)]
        [bool]
        $DefaultYes = $true
    )

    $suffix = if ($DefaultYes) { '[Y/n]' } else { '[y/N]' }
    $answer = Read-Host -Prompt "$Prompt $suffix"
    if ([string]::IsNullOrWhiteSpace($answer)) {
        return $DefaultYes
    }

    return ($answer -match '^(y|yes)$')
}

function Invoke-ModulePackageOperation {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ModuleName,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Add', 'Update')]
        [string]
        $Action,

        [Parameter(Mandatory = $false)]
        [object]
        $ExistingPackage,

        [Parameter(Mandatory = $true)]
        [string]
        $BasePath,

        [Parameter(Mandatory = $false)]
        [hashtable]
        $Versions,

        [Parameter(Mandatory = $false)]
        [bool]
        $PromptBeforeUpdate = $false,

        [Parameter(Mandatory = $false)]
        [bool]
        $ForcePromptBypass = $false,

        [Parameter(Mandatory = $false)]
        [bool]
        $SkipVersionCheck = $false,

        [Parameter(Mandatory = $false)]
        [bool]
        $WaitForCompletion = $false,

        [Parameter(Mandatory = $false)]
        [int]
        $TimeoutMinutes = 30,

        [Parameter(Mandatory = $false)]
        [int]
        $PollSeconds = 15,

        [Parameter(Mandatory = $false)]
        [bool]
        $FailHardOnAuthFailure = $true
    )

    $existingState = if ($ExistingPackage) { [string]$ExistingPackage.properties.provisioningState } else { $null }
    $existingVersion = if ($ExistingPackage) { [string]$ExistingPackage.properties.version } else { $null }

    $pinnedVersion = $null
    if ($Versions -and $Versions.ContainsKey($ModuleName)) {
        $pinnedVersion = [string]$Versions[$ModuleName]
    }

    $galleryInfo = Get-PSGalleryModuleVersion -ModuleName $ModuleName -PinnedVersion $pinnedVersion
    $targetVersion = if ($galleryInfo.Resolved) { $galleryInfo.Version } else { $null }

    # Prefer a versioned content URI when we know the target version (deterministic import).
    $contentUri = Get-GalleryContentUri -ModuleName $ModuleName -Version $targetVersion

    # Version precheck: skip update when installed is already at/above the gallery (or pin) target.
    if ($Action -eq 'Update' -and -not $SkipVersionCheck) {
        if (-not $galleryInfo.Resolved) {
            Write-Warning "Could not resolve gallery version for '$ModuleName' ($($galleryInfo.Error)). Will attempt import without a version precheck."
        }
        elseif ([string]::IsNullOrWhiteSpace($existingVersion)) {
            Write-Host "Precheck '$ModuleName': installed version unknown; import will be attempted (gallery $($galleryInfo.Source)=$targetVersion)."
        }
        else {
            $comparison = Compare-PackageVersionString -Left $existingVersion -Right $targetVersion
            if ($comparison -eq 0) {
                Write-Host "Precheck '$ModuleName': already current (installed=$existingVersion, target=$targetVersion). Skipping."
                return [PSCustomObject]@{
                    ModuleName        = $ModuleName
                    Action            = 'UpToDate'
                    ProvisioningState = $existingState
                    Version           = $existingVersion
                    TargetVersion     = $targetVersion
                    ContentUri        = $contentUri
                    Succeeded         = $true
                    ErrorMessage      = $null
                }
            }

            if ($comparison -eq 1) {
                Write-Host "Precheck '$ModuleName': installed version $existingVersion is newer than target $targetVersion. Skipping."
                return [PSCustomObject]@{
                    ModuleName        = $ModuleName
                    Action            = 'UpToDate'
                    ProvisioningState = $existingState
                    Version           = $existingVersion
                    TargetVersion     = $targetVersion
                    ContentUri        = $contentUri
                    Succeeded         = $true
                    ErrorMessage      = $null
                }
            }

            if ($comparison -eq -1) {
                Write-Host "Precheck '$ModuleName': update available (installed=$existingVersion → target=$targetVersion)."
            }
            else {
                Write-Warning "Precheck '$ModuleName': could not compare installed='$existingVersion' to target='$targetVersion'. Import will be attempted."
            }
        }
    }
    elseif ($Action -eq 'Add' -and $galleryInfo.Resolved) {
        Write-Host "Precheck '$ModuleName': not installed; will add gallery version $targetVersion."
    }

    if ($PromptBeforeUpdate -and -not $ForcePromptBypass) {
        $versionLabel = if ($existingVersion) { $existingVersion } else { 'unknown' }
        $targetLabel = if ($targetVersion) { $targetVersion } else { 'latest' }
        $stateLabel = if ($existingState) { $existingState } else { 'n/a' }
        $prompt = "Update package '$ModuleName' (installed=$versionLabel → target=$targetLabel, state=$stateLabel)?"
        if (-not (Read-YesNo -Prompt $prompt -DefaultYes $true)) {
            Write-Host "Skipped '$ModuleName' (user declined)."
            return [PSCustomObject]@{
                ModuleName        = $ModuleName
                Action            = 'Declined'
                ProvisioningState = $existingState
                Version           = $existingVersion
                TargetVersion     = $targetVersion
                ContentUri        = $contentUri
                Succeeded         = $true
                ErrorMessage      = $null
            }
        }
    }

    Write-Host "$Action package '$ModuleName' from $contentUri"

    # Nested helpers do not get the parent $PSCmdlet; honor common WhatIf preference instead.
    if ($WhatIfPreference) {
        Write-Host "  WhatIf: would $Action package '$ModuleName'."
        return [PSCustomObject]@{
            ModuleName        = $ModuleName
            Action            = 'WhatIf'
            ProvisioningState = $existingState
            Version           = $existingVersion
            TargetVersion     = $targetVersion
            ContentUri        = $contentUri
            Succeeded         = $true
            ErrorMessage      = $null
        }
    }

    try {
        $null = Set-RuntimeEnvironmentPackage -BasePath $BasePath -PackageName $ModuleName -ContentUri $contentUri
    }
    catch {
        Write-Error "Failed to $Action package '$ModuleName': $_"
        return [PSCustomObject]@{
            ModuleName        = $ModuleName
            Action            = $Action
            ProvisioningState = 'Failed'
            Version           = $null
            TargetVersion     = $targetVersion
            ContentUri        = $contentUri
            Succeeded         = $false
            ErrorMessage      = $_.Exception.Message
        }
    }

    if ($WaitForCompletion) {
        Write-Host "  Waiting for import of '$ModuleName'..."
        $waitResult = Wait-PackageProvisioning `
            -BasePath $BasePath `
            -PackageName $ModuleName `
            -TimeoutMinutes $TimeoutMinutes `
            -PollSeconds $PollSeconds

        if (-not $waitResult.Succeeded) {
            Write-Warning "Package '$ModuleName' did not succeed: $($waitResult.ErrorMessage)"
            if ($FailHardOnAuthFailure -and $ModuleName -eq $authName) {
                throw "Microsoft.Graph.Authentication failed to import. Dependent modules were not processed. $($waitResult.ErrorMessage)"
            }
        }

        return [PSCustomObject]@{
            ModuleName        = $ModuleName
            Action            = $Action
            ProvisioningState = $waitResult.ProvisioningState
            Version           = $waitResult.Version
            TargetVersion     = $targetVersion
            ContentUri        = $contentUri
            Succeeded         = $waitResult.Succeeded
            ErrorMessage      = $waitResult.ErrorMessage
        }
    }

    return [PSCustomObject]@{
        ModuleName        = $ModuleName
        Action            = $Action
        ProvisioningState = 'Submitted'
        Version           = $null
        TargetVersion     = $targetVersion
        ContentUri        = $contentUri
        Succeeded         = $true
        ErrorMessage      = $null
    }
}

# --- Main ---

$context = Get-RequiredAzContext
if (-not $SubscriptionId) {
    $SubscriptionId = $context.Subscription.Id
}

if (-not $SubscriptionId) {
    throw 'Subscription ID could not be resolved. Pass -SubscriptionId or select a subscription with Set-AzContext.'
}

$basePath = "/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroupName/providers/Microsoft.Automation/automationAccounts/$AutomationAccountName/runtimeEnvironments/$RuntimeEnvironmentName"

Write-Host "Validating runtime environment '$RuntimeEnvironmentName'..."
try {
    $null = Invoke-AutomationRest -Method GET -RelativePath $basePath
}
catch {
    throw "Runtime environment '$RuntimeEnvironmentName' was not found on automation account '$AutomationAccountName' in resource group '$ResourceGroupName'. $_"
}

Write-Host 'Listing existing packages on the runtime environment...'
$existingPackages = @(Get-RuntimeEnvironmentPackages -BasePath $basePath)
$existingByName = @{}
foreach ($pkg in $existingPackages) {
    $existingByName[$pkg.name] = $pkg
}

Write-Host "Found $($existingPackages.Count) existing package(s)."

$results = [System.Collections.Generic.List[object]]::new()

if ($PSCmdlet.ParameterSetName -eq 'UpdateExisting') {
    Write-Host 'Mode: Update existing modules'

    $updatablePackages = @($existingPackages | Where-Object {
            # Platform default packages (e.g. built-in Az) are not re-imported from the gallery.
            -not [bool]($_.properties.default)
        })

    $defaultSkipped = @($existingPackages | Where-Object { [bool]($_.properties.default) })
    if ($defaultSkipped.Count -gt 0) {
        Write-Host "Skipping $($defaultSkipped.Count) platform default package(s): $($defaultSkipped.name -join ', ')"
    }

    if ($UpdateModuleNames -and $UpdateModuleNames.Count -gt 0) {
        $candidates = [System.Collections.Generic.List[string]]::new()
        foreach ($name in $UpdateModuleNames) {
            if ([string]::IsNullOrWhiteSpace($name)) {
                continue
            }

            $match = $updatablePackages | Where-Object { $_.name -eq $name } | Select-Object -First 1
            if (-not $match) {
                $onRuntime = $existingByName.ContainsKey($name)
                if ($onRuntime -and [bool]$existingByName[$name].properties.default) {
                    Write-Warning "Package '$name' is a platform default package and cannot be updated via gallery import. Skipping."
                }
                elseif (-not $onRuntime) {
                    Write-Warning "Package '$name' is not installed on runtime '$RuntimeEnvironmentName'. Skipping (update-existing does not add new modules)."
                }
                else {
                    Write-Warning "Package '$name' was not selected for update. Skipping."
                }
                continue
            }

            if (-not $candidates.Contains($name)) {
                $candidates.Add($name)
            }
        }

        $orderedModules = Get-OrderedModuleNames -Names @($candidates)
        Write-Host "Updating specified existing modules: $($orderedModules -join ', ')"
    }
    else {
        $orderedModules = Get-OrderedModuleNames -Names @($updatablePackages | ForEach-Object { $_.name })
        Write-Host "Updating ALL non-default existing modules ($($orderedModules.Count)): $($orderedModules -join ', ')"
    }

    if ($orderedModules.Count -eq 0) {
        Write-Warning 'No existing packages matched for update.'
        return @()
    }

    if (-not $Force) {
        Write-Host 'You will be prompted for each package. Use -Force to update without prompts.'
    }

    Write-Host ''

    foreach ($moduleName in $orderedModules) {
        $result = Invoke-ModulePackageOperation `
            -ModuleName $moduleName `
            -Action 'Update' `
            -ExistingPackage $existingByName[$moduleName] `
            -BasePath $basePath `
            -Versions $ModuleVersions `
            -PromptBeforeUpdate $true `
            -ForcePromptBypass:$Force.IsPresent `
            -SkipVersionCheck:$SkipVersionCheck.IsPresent `
            -WaitForCompletion:$Wait.IsPresent `
            -TimeoutMinutes $WaitTimeoutMinutes `
            -PollSeconds $PollIntervalSeconds `
            -FailHardOnAuthFailure $false

        $results.Add($result)
    }
}
else {
    Write-Host 'Mode: Ensure required modules'

    if (-not $ModuleNames -or $ModuleNames.Count -eq 0) {
        throw 'ModuleNames cannot be empty.'
    }

    $orderedModules = Get-OrderedModuleNames -Names $ModuleNames
    Write-Host "Modules to process: $($orderedModules -join ', ')"
    Write-Host ''

    foreach ($moduleName in $orderedModules) {
        $existing = $existingByName[$moduleName]
        $exists = $null -ne $existing
        $existingState = if ($exists) { [string]$existing.properties.provisioningState } else { $null }
        $existingVersion = if ($exists) { [string]$existing.properties.version } else { $null }

        if ($SkipExisting -and $exists -and $existingState -ne 'Failed') {
            Write-Host "Skipping existing package '$moduleName' (state=$existingState version=$existingVersion) because -SkipExisting was specified."
            $results.Add([PSCustomObject]@{
                    ModuleName        = $moduleName
                    Action            = 'Skipped'
                    ProvisioningState = $existingState
                    Version           = $existingVersion
                    TargetVersion     = $null
                    ContentUri        = $null
                    Succeeded         = $true
                    ErrorMessage      = $null
                })
            continue
        }

        $action = if ($exists) { 'Update' } else { 'Add' }

        $result = Invoke-ModulePackageOperation `
            -ModuleName $moduleName `
            -Action $action `
            -ExistingPackage $existing `
            -BasePath $basePath `
            -Versions $ModuleVersions `
            -PromptBeforeUpdate $false `
            -ForcePromptBypass $true `
            -SkipVersionCheck:$SkipVersionCheck.IsPresent `
            -WaitForCompletion:$Wait.IsPresent `
            -TimeoutMinutes $WaitTimeoutMinutes `
            -PollSeconds $PollIntervalSeconds `
            -FailHardOnAuthFailure $true

        $results.Add($result)
    }
}

Write-Host ''
Write-Host 'Summary:'
$results | Format-Table -AutoSize ModuleName, Action, Version, TargetVersion, ProvisioningState, Succeeded | Out-Host

$failed = @($results | Where-Object { -not $_.Succeeded })
if ($failed.Count -gt 0) {
    throw "$($failed.Count) module operation(s) failed: $($failed.ModuleName -join ', ')"
}

if (-not $Wait) {
    Write-Host "Imports were submitted asynchronously. Check the portal (Runtime environments → $RuntimeEnvironmentName → Packages) or re-run with -Wait."
}
else {
    Write-Host 'All requested module operations completed successfully (or were skipped/declined).'
}

return $results
