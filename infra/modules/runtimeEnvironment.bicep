@description('Name of the parent Automation Account.')
param automationAccountName string

@description('Azure region (must match the Automation Account).')
param location string

@description('Name of the custom PowerShell 7.2 runtime environment.')
@minLength(1)
param runtimeEnvironmentName string = 'EAM-PS72-Graph'

@description('Human-readable description for the runtime environment.')
param runtimeDescription string = 'PowerShell 7.2 runtime for EAM-AutoUpdater with Microsoft Graph modules.'

@description('Resource tags applied to the runtime environment.')
param tags object = {}

@description('''
Microsoft Graph PowerShell modules to import from the PowerShell Gallery into this runtime.
Import order matters: Microsoft.Graph.Authentication must be first.
''')
param graphModuleNames array = [
  'Microsoft.Graph.Authentication'
  'Microsoft.Graph.Beta.DeviceManagement.Actions'
  'Microsoft.Graph.Beta.Devices.CorporateManagement'
  'Microsoft.Graph.Groups'
  'Microsoft.Graph.Beta.DeviceManagement'
]

@description('Optional module version map (module name -> version). Empty string means latest from the gallery.')
param graphModuleVersions object = {}

resource automationAccount 'Microsoft.Automation/automationAccounts@2024-10-23' existing = {
  name: automationAccountName
}

resource runtimeEnvironment 'Microsoft.Automation/automationAccounts/runtimeEnvironments@2024-10-23' = {
  parent: automationAccount
  name: runtimeEnvironmentName
  location: location
  tags: tags
  properties: {
    description: runtimeDescription
    runtime: {
      language: 'PowerShell'
      version: '7.2'
    }
  }
}

// Authentication must import before dependent Graph modules.
var authenticationModuleName = 'Microsoft.Graph.Authentication'
var authenticationVersion = graphModuleVersions[?authenticationModuleName] ?? ''
var authenticationContentUri = empty(authenticationVersion)
  ? 'https://www.powershellgallery.com/api/v2/package/${authenticationModuleName}'
  : 'https://www.powershellgallery.com/api/v2/package/${authenticationModuleName}/${authenticationVersion}'

resource authenticationPackage 'Microsoft.Automation/automationAccounts/runtimeEnvironments/packages@2024-10-23' = {
  parent: runtimeEnvironment
  name: authenticationModuleName
  properties: {
    contentLink: {
      uri: authenticationContentUri
    }
  }
}

var dependentModules = filter(graphModuleNames, moduleName => moduleName != authenticationModuleName)

resource graphPackages 'Microsoft.Automation/automationAccounts/runtimeEnvironments/packages@2024-10-23' = [
  for moduleName in dependentModules: {
    parent: runtimeEnvironment
    name: moduleName
    dependsOn: [
      authenticationPackage
    ]
    properties: {
      contentLink: {
        // Pin via graphModuleVersions when set; otherwise use latest gallery package.
        uri: contains(graphModuleVersions, moduleName) && !empty(graphModuleVersions[moduleName])
          ? 'https://www.powershellgallery.com/api/v2/package/${moduleName}/${graphModuleVersions[moduleName]}'
          : 'https://www.powershellgallery.com/api/v2/package/${moduleName}'
      }
    }
  }
]

@description('Runtime environment name.')
output name string = runtimeEnvironment.name

@description('Runtime environment resource ID.')
output id string = runtimeEnvironment.id

@description('Modules requested for import into the runtime environment.')
output moduleNames array = graphModuleNames
