#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Applications

<#
.SYNOPSIS
    Assigns Microsoft Graph application permissions to an Azure Automation managed identity.

.DESCRIPTION
    Grants the application app roles required by EAM-AutoUpdater (Invoke-EAMAutoUpdate.ps1)
    to a system-assigned managed identity. Mirrors the permission set documented in
    Documentation/02-Setup-AzureAutomationAccount.md and the repository README.

.PARAMETER ManagedIdentityObjectId
    Object (principal) ID of the Automation Account system-assigned managed identity.
    Available as the managedIdentityPrincipalId output from the Bicep deployment.

.PARAMETER IncludeEspPermission
    When set, also assigns DeviceManagementServiceConfig.ReadWrite.All for Enrollment
    Status Page updates (-UpdateESP). Omit this switch if you do not update ESP profiles.

.PARAMETER WhatIf
    Shows which app roles would be assigned without making changes.

.EXAMPLE
    .\Assign-GraphAppRoles.ps1 -ManagedIdentityObjectId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'

.EXAMPLE
    .\Assign-GraphAppRoles.ps1 -ManagedIdentityObjectId $principalId -IncludeEspPermission
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $ManagedIdentityObjectId,

    [Parameter(Mandatory = $false)]
    [switch]
    $IncludeEspPermission
)

$ErrorActionPreference = 'Stop'

$requiredPermissions = @(
    'DeviceManagementManagedDevices.Read.All'
    'DeviceManagementConfiguration.Read.All'
    'DeviceManagementApps.ReadWrite.All'
    'Group.Read.All'
    'DeviceManagementRBAC.Read.All'
)

if ($IncludeEspPermission) {
    $requiredPermissions += 'DeviceManagementServiceConfig.ReadWrite.All'
}

Write-Host "Connecting to Microsoft Graph (AppRoleAssignment.ReadWrite.All, Application.Read.All)..."
Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All', 'Application.Read.All' -NoWelcome | Out-Null

$graphAppId = '00000003-0000-0000-c000-000000000000'
$graphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"
if (-not $graphServicePrincipal) {
    throw "Microsoft Graph service principal ($graphAppId) was not found in this tenant."
}

$existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityObjectId -All
$existingRoleIds = @($existingAssignments | ForEach-Object { $_.AppRoleId })

$rolesToAssign = $graphServicePrincipal.AppRoles |
    Where-Object {
        $_.Value -in $requiredPermissions -and
        $_.AllowedMemberTypes -contains 'Application'
    }

$missing = $requiredPermissions | Where-Object { $_ -notin @($rolesToAssign.Value) }
if ($missing.Count -gt 0) {
    throw "The following app roles were not found on Microsoft Graph: $($missing -join ', ')"
}

foreach ($role in $rolesToAssign) {
    if ($existingRoleIds -contains $role.Id) {
        Write-Host "Already assigned: $($role.Value)"
        continue
    }

    $assignment = @{
        ServicePrincipalId = $ManagedIdentityObjectId
        PrincipalId        = $ManagedIdentityObjectId
        ResourceId         = $graphServicePrincipal.Id
        AppRoleId          = $role.Id
    }

    if ($PSCmdlet.ShouldProcess($role.Value, 'Assign Graph application app role to managed identity')) {
        New-MgServicePrincipalAppRoleAssignment @assignment | Out-Null
        Write-Host "Assigned: $($role.Value)"
    }
}

Write-Host ''
Write-Host 'Done. Verify permissions on the managed identity Enterprise application in Entra ID (Permissions blade).'
Write-Host 'If Intune Multi Admin Approval is enabled for Applications, exclude this managed identity from MAA.'
