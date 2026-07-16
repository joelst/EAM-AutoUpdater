targetScope = 'resourceGroup'

// ---------------------------------------------------------------------------
// EAM-AutoUpdater — Azure Automation host (new or existing) + PS 7.2 runtime
// ---------------------------------------------------------------------------

@description('Create a new Automation Account or reuse an existing one.')
@allowed([
  'New'
  'Existing'
])
param automationAccountMode string = 'New'

@description('Automation Account name (created or existing).')
param automationAccountName string

@description('''
Resource group that contains the existing Automation Account.
Only used when automationAccountMode is Existing. Defaults to the deployment resource group.
''')
param existingAutomationAccountResourceGroup string = resourceGroup().name

@description('Azure region for new resources. Defaults to the deployment resource group location.')
param location string = resourceGroup().location

@description('Resource tags applied to created resources.')
param tags object = {
  application: 'EAM-AutoUpdater'
}

@description('Allow public network access on non-ARM endpoints when creating a new account.')
param publicNetworkAccess bool = true

@description('SKU for a newly created Automation Account.')
@allowed([
  'Basic'
  'Free'
])
param skuName string = 'Basic'

@description('Name of the custom PowerShell 7.2 runtime environment.')
param runtimeEnvironmentName string = 'EAM-PS72-Graph'

@description('Microsoft Graph modules to import into the runtime environment (Authentication first).')
param graphModuleNames array = [
  'Microsoft.Graph.Authentication'
  'Microsoft.Graph.Beta.DeviceManagement.Actions'
  'Microsoft.Graph.Beta.Devices.CorporateManagement'
  'Microsoft.Graph.Groups'
  'Microsoft.Graph.Beta.DeviceManagement'
]

@description('Optional module version pins (module name -> version). Empty object uses latest gallery versions.')
param graphModuleVersions object = {}

@description('Deploy the Invoke-EAMAutoUpdate runbook shell (content published separately).')
param deployRunbook bool = true

@description('Runbook name when deployRunbook is true.')
param runbookName string = 'Invoke-EAMAutoUpdate'

@description('Deploy a schedule and link it to the runbook. Requires deployRunbook = true.')
param deploySchedule bool = false

@description('Schedule resource name.')
param scheduleName string = 'EAM-Daily'

@description('Schedule frequency when deploySchedule is true.')
@allowed([
  'Day'
  'Hour'
  'Week'
  'Month'
])
param scheduleFrequency string = 'Day'

@description('Schedule interval (e.g. 1 day). Keep spacing >= 1 hour for EAM report refresh.')
param scheduleInterval int = 1

@description('First schedule run (ISO 8601). Required when deploySchedule is true; must be in the future.')
param scheduleStartTime string = ''

@description('Optional schedule end time (ISO 8601).')
param scheduleExpiryTime string = ''

@description('Time zone for the schedule.')
param scheduleTimeZone string = 'UTC'

// ----- Automation Account (create or reference) -----

module automationAccountNew 'modules/automationAccount.bicep' = if (automationAccountMode == 'New') {
  name: 'automation-account-new'
  params: {
    name: automationAccountName
    location: location
    tags: tags
    publicNetworkAccess: publicNetworkAccess
    skuName: skuName
  }
}

module automationAccountExisting 'modules/automationAccountReference.bicep' = if (automationAccountMode == 'Existing') {
  name: 'automation-account-existing'
  scope: resourceGroup(existingAutomationAccountResourceGroup)
  params: {
    name: automationAccountName
  }
}

var resolvedAutomationAccountName = automationAccountMode == 'New'
  ? automationAccountNew!.outputs.name
  : automationAccountExisting!.outputs.name

var resolvedAutomationAccountId = automationAccountMode == 'New'
  ? automationAccountNew!.outputs.id
  : automationAccountExisting!.outputs.id

var resolvedLocation = automationAccountMode == 'New'
  ? automationAccountNew!.outputs.location
  : automationAccountExisting!.outputs.location

var resolvedPrincipalId = automationAccountMode == 'New'
  ? automationAccountNew!.outputs.principalId
  : automationAccountExisting!.outputs.principalId

// Runtime environment + packages are always deployed into the same RG as the AA.
// When the existing AA is in another RG, deploy child resources into that RG.
module runtimeEnvironment 'modules/runtimeEnvironment.bicep' = {
  name: 'eam-runtime-environment'
  scope: resourceGroup(automationAccountMode == 'Existing'
    ? existingAutomationAccountResourceGroup
    : resourceGroup().name)
  params: {
    automationAccountName: resolvedAutomationAccountName
    location: resolvedLocation
    runtimeEnvironmentName: runtimeEnvironmentName
    tags: tags
    graphModuleNames: graphModuleNames
    graphModuleVersions: graphModuleVersions
  }
}

module runbook 'modules/runbook.bicep' = if (deployRunbook) {
  name: 'eam-runbook'
  scope: resourceGroup(automationAccountMode == 'Existing'
    ? existingAutomationAccountResourceGroup
    : resourceGroup().name)
  params: {
    automationAccountName: resolvedAutomationAccountName
    location: resolvedLocation
    runbookName: runbookName
    runtimeEnvironmentName: runtimeEnvironment.outputs.name
    tags: tags
  }
}

module schedule 'modules/schedule.bicep' = if (deployRunbook && deploySchedule) {
  name: 'eam-schedule'
  scope: resourceGroup(automationAccountMode == 'Existing'
    ? existingAutomationAccountResourceGroup
    : resourceGroup().name)
  params: {
    automationAccountName: resolvedAutomationAccountName
    scheduleName: scheduleName
    runbookName: runbookName
    frequency: scheduleFrequency
    interval: scheduleInterval
    startTime: scheduleStartTime
    expiryTime: scheduleExpiryTime
    timeZone: scheduleTimeZone
  }
  dependsOn: [
    runbook
  ]
}

// ----- Outputs -----

@description('Deployment mode used for the Automation Account.')
output automationAccountMode string = automationAccountMode

@description('Automation Account name.')
output automationAccountName string = resolvedAutomationAccountName

@description('Automation Account resource ID.')
output automationAccountId string = resolvedAutomationAccountId

@description('System-assigned managed identity principal (object) ID. Required for Graph app-role assignment.')
output managedIdentityPrincipalId string = resolvedPrincipalId

@description('Runtime environment name.')
output runtimeEnvironmentName string = runtimeEnvironment.outputs.name

@description('Runtime environment resource ID.')
output runtimeEnvironmentId string = runtimeEnvironment.outputs.id

@description('Runbook name when deployRunbook is true; otherwise empty.')
output runbookName string = deployRunbook ? runbook!.outputs.name : ''

@description('Next step: assign Microsoft Graph application permissions to the managed identity.')
output graphPermissionScriptHint string = empty(resolvedPrincipalId)
  ? 'System-assigned managed identity principal ID was empty. Enable system-assigned identity on the Automation Account, then re-run deployment or assign permissions manually.'
  : 'pwsh -File infra/scripts/Assign-GraphAppRoles.ps1 -ManagedIdentityObjectId ${resolvedPrincipalId}'
