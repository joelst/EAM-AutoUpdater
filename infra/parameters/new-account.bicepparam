using '../main.bicep'

// Create a new Automation Account with the EAM PowerShell 7.2 runtime environment.

param automationAccountMode = 'New'
param automationAccountName = 'aa-eam-autoupdater' // must be unique in the tenant
param location = 'westeurope'

param tags = {
  application: 'EAM-AutoUpdater'
  environment: 'prod'
}

param publicNetworkAccess = true
param runtimeEnvironmentName = 'EAM-PS72-Graph'

// Deploy runbook shell (publish content with infra/scripts/Publish-EAMRunbook.ps1)
param deployRunbook = true
param runbookName = 'Invoke-EAMAutoUpdate'

// Optional daily schedule (startTime must be in the future when enabled)
param deploySchedule = false
// param scheduleStartTime = '2026-07-15T06:00:00+00:00'
// param scheduleTimeZone = 'W. Europe Standard Time'
// param scheduleFrequency = 'Day'
// param scheduleInterval = 1
