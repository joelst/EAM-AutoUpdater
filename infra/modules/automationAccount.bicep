@description('Name of the Azure Automation Account to create.')
param name string

@description('Azure region for the Automation Account.')
param location string = resourceGroup().location

@description('Resource tags applied to the Automation Account.')
param tags object = {}

@description('Allow public network access on non-ARM endpoints (webhooks/agent).')
param publicNetworkAccess bool = true

@description('SKU name for the Automation Account.')
@allowed([
  'Basic'
  'Free'
])
param skuName string = 'Basic'

resource automationAccount 'Microsoft.Automation/automationAccounts@2024-10-23' = {
  name: name
  location: location
  tags: tags
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    sku: {
      name: skuName
    }
    publicNetworkAccess: publicNetworkAccess
    encryption: {
      keySource: 'Microsoft.Automation'
    }
  }
}

@description('Automation Account name.')
output name string = automationAccount.name

@description('Automation Account resource ID.')
output id string = automationAccount.id

@description('Azure region of the Automation Account.')
output location string = automationAccount.location

@description('System-assigned managed identity principal (object) ID.')
output principalId string = automationAccount.identity.principalId

@description('System-assigned managed identity tenant ID.')
output tenantId string = automationAccount.identity.tenantId
