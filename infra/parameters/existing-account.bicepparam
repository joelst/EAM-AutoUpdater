using '../main.bicep'

// Attach the EAM PowerShell 7.2 runtime environment (and optional runbook) to an existing Automation Account.
// The account should already have a system-assigned managed identity enabled.

param automationAccountMode = 'Existing'
param automationAccountName = 'aa-existing-automation'
param existingAutomationAccountResourceGroup = 'rg-existing-automation'

param tags = {
  application: 'EAM-AutoUpdater'
}

param runtimeEnvironmentName = 'EAM-PS72-Graph'

param deployRunbook = true
param runbookName = 'Invoke-EAMAutoUpdate'

param deploySchedule = false
