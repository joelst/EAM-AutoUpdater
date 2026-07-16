@description('Name of an existing Azure Automation Account in the current resource group scope.')
param name string

resource automationAccount 'Microsoft.Automation/automationAccounts@2024-10-23' existing = {
  name: name
}

@description('Automation Account name.')
output name string = automationAccount.name

@description('Automation Account resource ID.')
output id string = automationAccount.id

@description('Azure region of the Automation Account.')
output location string = automationAccount.location

@description('System-assigned managed identity principal (object) ID. Empty if identity is not enabled.')
output principalId string = automationAccount.?identity.?principalId ?? ''

@description('System-assigned managed identity tenant ID. Empty if identity is not enabled.')
output tenantId string = automationAccount.?identity.?tenantId ?? ''
