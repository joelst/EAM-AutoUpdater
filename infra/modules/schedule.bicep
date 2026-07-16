@description('Name of the parent Automation Account.')
param automationAccountName string

@description('Name of the schedule resource.')
param scheduleName string = 'EAM-Daily'

@description('Name of the runbook to link to this schedule.')
param runbookName string

@description('Human-readable schedule description.')
param scheduleDescription string = 'Daily EAM-AutoUpdater execution. Keep at least 1 hour between runs so the Intune EAM report can refresh.'

@description('Schedule frequency.')
@allowed([
  'Day'
  'Hour'
  'Week'
  'Month'
])
param frequency string = 'Day'

@description('Interval for the frequency (e.g. 1 = every day when frequency is Day). Minimum practical spacing for EAM is 1 hour.')
param interval int = 1

@description('First run time (ISO 8601). Must be in the future relative to deployment.')
param startTime string

@description('Optional schedule end time (ISO 8601). Leave empty for no expiry.')
param expiryTime string = ''

@description('Time zone for the schedule (Windows or IANA id supported by Automation).')
param timeZone string = 'UTC'

resource automationAccount 'Microsoft.Automation/automationAccounts@2024-10-23' existing = {
  name: automationAccountName
}

var scheduleProperties = union(
  {
    description: scheduleDescription
    frequency: frequency
    interval: interval
    startTime: startTime
    timeZone: timeZone
  },
  empty(expiryTime) ? {} : { expiryTime: expiryTime }
)

resource schedule 'Microsoft.Automation/automationAccounts/schedules@2024-10-23' = {
  parent: automationAccount
  name: scheduleName
  properties: scheduleProperties
}

// jobSchedules name must be a GUID. Use a stable GUID derived from names so redeploys are idempotent.
var jobScheduleName = guid(automationAccountName, runbookName, scheduleName)

resource jobSchedule 'Microsoft.Automation/automationAccounts/jobSchedules@2024-10-23' = {
  parent: automationAccount
  name: jobScheduleName
  properties: {
    runbook: {
      name: runbookName
    }
    schedule: {
      name: schedule.name
    }
  }
}

@description('Schedule name.')
output name string = schedule.name

@description('Schedule resource ID.')
output id string = schedule.id

@description('Job schedule resource ID.')
output jobScheduleId string = jobSchedule.id
