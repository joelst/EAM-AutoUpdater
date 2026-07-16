#Requires -Version 7.0
#Requires -Modules Az.Accounts, Az.Automation

<#
.SYNOPSIS
    Publishes Invoke-EAMAutoUpdate.ps1 into an Azure Automation runbook.

.DESCRIPTION
    Imports the repository runbook script into a PowerShell runbook that was created
    by the Bicep deployment (draft shell linked to the EAM runtime environment).
    After import, uncomment and customize the sample Invoke-EAMAutoUpdate call at the
    bottom of the script before relying on scheduled runs.

.PARAMETER ResourceGroupName
    Resource group containing the Automation Account.

.PARAMETER AutomationAccountName
    Automation Account name.

.PARAMETER RunbookName
    Target runbook name. Defaults to Invoke-EAMAutoUpdate.

.PARAMETER ScriptPath
    Path to Invoke-EAMAutoUpdate.ps1. Defaults to the repo root copy next to infra/.

.PARAMETER Force
    Overwrite existing runbook content if the runbook already has published content.

.EXAMPLE
    .\Publish-EAMRunbook.ps1 -ResourceGroupName 'rg-eam' -AutomationAccountName 'aa-eam-prod'
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $ResourceGroupName,

    [Parameter(Mandatory = $true)]
    [string]
    $AutomationAccountName,

    [Parameter(Mandatory = $false)]
    [string]
    $RunbookName = 'Invoke-EAMAutoUpdate',

    [Parameter(Mandatory = $false)]
    [string]
    $ScriptPath = (Join-Path -Path (Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent) -ChildPath 'Invoke-EAMAutoUpdate.ps1'),

    [Parameter(Mandatory = $false)]
    [switch]
    $Force
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ScriptPath)) {
    throw "Runbook script not found at '$ScriptPath'."
}

$context = Get-AzContext
if (-not $context) {
    throw 'No Azure context. Run Connect-AzAccount first.'
}

Write-Host "Importing '$ScriptPath' into runbook '$RunbookName' on account '$AutomationAccountName'..."

$importParams = @{
    ResourceGroupName     = $ResourceGroupName
    AutomationAccountName = $AutomationAccountName
    Name                  = $RunbookName
    Type                  = 'PowerShell'
    Path                  = $ScriptPath
    Published             = $true
    Force                 = $Force.IsPresent
}

if ($PSCmdlet.ShouldProcess($RunbookName, 'Import and publish Automation runbook content')) {
    Import-AzAutomationRunbook @importParams | Out-Null
    Write-Host 'Publish completed.'
    Write-Host 'Next: edit the runbook in the portal (or re-import a customized copy) and uncomment the Invoke-EAMAutoUpdate sample call with your parameters.'
}
