@description('Name of the parent Automation Account.')
param automationAccountName string

@description('Azure region (must match the Automation Account).')
param location string

@description('Runbook name (letters, numbers, hyphens, underscores; must start with a letter).')
param runbookName string = 'Invoke-EAMAutoUpdate'

@description('Runtime environment name to associate with the runbook (PowerShell 7.2).')
param runtimeEnvironmentName string

@description('Human-readable runbook description.')
param runbookDescription string = 'EAM-AutoUpdater: publish Intune Enterprise Application Management catalog updates.'

@description('Resource tags applied to the runbook.')
param tags object = {}

@description('Enable progress logging for runbook jobs.')
param logProgress bool = true

@description('Enable verbose logging for runbook jobs.')
param logVerbose bool = false

resource automationAccount 'Microsoft.Automation/automationAccounts@2024-10-23' existing = {
  name: automationAccountName
}

// Creates the runbook shell linked to the EAM runtime environment.
// Content must be published separately (portal paste, or infra/scripts/Publish-EAMRunbook.ps1).
resource runbook 'Microsoft.Automation/automationAccounts/runbooks@2024-10-23' = {
  parent: automationAccount
  name: runbookName
  location: location
  tags: tags
  properties: {
    runbookType: 'PowerShell'
    runtimeEnvironment: runtimeEnvironmentName
    description: runbookDescription
    logProgress: logProgress
    logVerbose: logVerbose
    draft: {}
  }
}

@description('Runbook name.')
output name string = runbook.name

@description('Runbook resource ID.')
output id string = runbook.id
