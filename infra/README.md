# EAM-AutoUpdater — Azure Automation (Bicep)

Deploys or reuses an **Azure Automation Account**, creates a **PowerShell 7.2 runtime environment** with the Microsoft Graph modules required by `Invoke-EAMAutoUpdate.ps1`, and optionally creates the runbook shell and a schedule.

## What gets deployed

| Resource | New account | Existing account |
| --- | --- | --- |
| Automation Account + system-assigned managed identity | Created | Referenced (not recreated) |
| Runtime environment `EAM-PS72-Graph` (PowerShell 7.2) | Created | Created on the existing account |
| Graph module packages (gallery) | Imported into the runtime | Imported into the runtime |
| Runbook shell `Invoke-EAMAutoUpdate` | Optional (`deployRunbook`) | Optional |
| Daily schedule + job link | Optional (`deploySchedule`) | Optional |
| Microsoft Graph app roles on the MI | Post-deploy script | Post-deploy script |

### Modules imported into the runtime

1. `Microsoft.Graph.Authentication` (first)
2. `Microsoft.Graph.Beta.DeviceManagement.Actions`
3. `Microsoft.Graph.Beta.Devices.CorporateManagement`
4. `Microsoft.Graph.Groups`
5. `Microsoft.Graph.Beta.DeviceManagement`

Gallery imports are asynchronous. After deployment, wait until packages show **Succeeded** / **Available** in the portal (**Automation Account → Runtime environments**) before testing the runbook.

## Prerequisites

- Azure CLI with Bicep (`az bicep version`) or Azure PowerShell
- Rights to deploy to the target resource group (**Contributor** is typical)
- For Graph permission assignment: **Application Administrator** (or equivalent) plus `Microsoft.Graph.Applications` PowerShell module
- Existing-account mode: Automation Account already exists; **system-assigned managed identity** should be enabled

## Quick start

### 1. Create a resource group (new account)

```bash
az group create --name rg-eam-autoupdater --location westeurope
```

### 2. Review parameters

- New account: [`parameters/new-account.bicepparam`](parameters/new-account.bicepparam)
- Existing account: [`parameters/existing-account.bicepparam`](parameters/existing-account.bicepparam)

Set a unique `automationAccountName` (global uniqueness rules apply for new accounts).

### 3. Validate and deploy

```bash
# From the repository root
az bicep build --file infra/main.bicep

az deployment group create \
  --resource-group rg-eam-autoupdater \
  --template-file infra/main.bicep \
  --parameters infra/parameters/new-account.bicepparam
```

Existing account example:

```bash
az deployment group create \
  --resource-group rg-eam-autoupdater \
  --template-file infra/main.bicep \
  --parameters infra/parameters/existing-account.bicepparam
```

> When `automationAccountMode` is `Existing` and the account lives in another resource group, child resources (runtime, runbook, schedule) are deployed into **that** resource group via module scope. The deployment command’s `--resource-group` is still required as the deployment scope; set `existingAutomationAccountResourceGroup` to the account’s RG.

### 4. Assign Microsoft Graph permissions

```powershell
$principalId = az deployment group show `
  --resource-group rg-eam-autoupdater `
  --name <deploymentName> `
  --query properties.outputs.managedIdentityPrincipalId.value -o tsv

# Include -IncludeEspPermission if you will run with -UpdateESP
pwsh -File infra/scripts/Assign-GraphAppRoles.ps1 -ManagedIdentityObjectId $principalId
```

### 5. Add or update runtime modules (optional / maintenance)

After deployment, or later when Graph modules need refreshing, import or update packages on the runtime environment:

```powershell
# Prompts for Connect-AzAccount when no valid Azure context is present
pwsh -File infra/scripts/Update-RuntimeEnvironmentModules.ps1 `
  -ResourceGroupName rg-eam-autoupdater `
  -AutomationAccountName aa-eam-autoupdater `
  -RuntimeEnvironmentName EAM-PS72-Graph `
  -Wait
```

- **Ensure required (default):** installs/updates the EAM Graph module list (`Microsoft.Graph.Authentication` first). Use `-SkipExisting` to only add missing modules.
- **Update existing:** `-UpdateExistingModules` re-imports packages already on the runtime:
  - Omit `-UpdateModuleNames` → **all** non-default installed packages (prompts **per module** unless `-Force`).
  - `-UpdateModuleNames 'A','B'` → only those packages if present (also prompts per module unless `-Force`).
- **Version precheck:** before updating an installed package, queries the PowerShell Gallery and **skips** (no prompt, no import) when the installed version is already current. Use `-SkipVersionCheck` to force re-import anyway.
- Use `-ModuleVersions @{ 'Microsoft.Graph.Authentication' = '2.26.1' }` to pin versions.
- Platform default packages (e.g. built-in Az on some REs) are skipped in update-existing mode.

### 6. Publish runbook content

The Bicep template creates an empty draft runbook linked to the EAM runtime. Publish the script from the repo:

```powershell
Connect-AzAccount
pwsh -File infra/scripts/Publish-EAMRunbook.ps1 `
  -ResourceGroupName rg-eam-autoupdater `
  -AutomationAccountName aa-eam-autoupdater `
  -Force
```

Then edit the runbook: uncomment and customize the sample `Invoke-EAMAutoUpdate ...` call (Teams webhook, `-UpdateESP`, exclusions, update rings, etc.).

### 7. Optional schedule

Set in the parameter file:

```bicep
param deploySchedule = true
param scheduleStartTime = '2026-07-15T06:00:00+00:00'  // must be in the future
param scheduleTimeZone = 'W. Europe Standard Time'
param scheduleFrequency = 'Day'
param scheduleInterval = 1
```

Keep at least **one hour** between runs so the Intune EAM report can refresh.

## Parameters (main)

| Parameter | Default | Description |
| --- | --- | --- |
| `automationAccountMode` | `New` | `New` or `Existing` |
| `automationAccountName` | *(required)* | Account name |
| `existingAutomationAccountResourceGroup` | deployment RG | RG of existing account |
| `location` | RG location | Region for new resources |
| `tags` | `application=EAM-AutoUpdater` | Tags |
| `publicNetworkAccess` | `true` | New account networking |
| `skuName` | `Basic` | New account SKU |
| `runtimeEnvironmentName` | `EAM-PS72-Graph` | Custom RE name |
| `graphModuleNames` | five Graph modules | Packages to import |
| `graphModuleVersions` | `{}` | Optional version pins |
| `deployRunbook` | `true` | Create runbook shell |
| `runbookName` | `Invoke-EAMAutoUpdate` | Runbook name |
| `deploySchedule` | `false` | Create schedule + job link |
| `scheduleStartTime` | `''` | Required when scheduling |

## Outputs

| Output | Description |
| --- | --- |
| `automationAccountName` / `automationAccountId` | Account identity |
| `managedIdentityPrincipalId` | Object ID for Graph app-role assignment |
| `runtimeEnvironmentName` / `runtimeEnvironmentId` | Custom RE |
| `runbookName` | When runbook deployment is enabled |
| `graphPermissionScriptHint` | Suggested permission script command |

## Layout

```
infra/
  main.bicep
  modules/
    automationAccount.bicep           # create AA + system-assigned MI
    automationAccountReference.bicep  # existing AA lookup
    runtimeEnvironment.bicep          # PS 7.2 RE + Graph packages
    runbook.bicep                     # runbook shell + RE link
    schedule.bicep                    # schedule + jobSchedules
  parameters/
    new-account.bicepparam
    existing-account.bicepparam
  scripts/
    Assign-GraphAppRoles.ps1
    Publish-EAMRunbook.ps1
    Update-RuntimeEnvironmentModules.ps1
  README.md
```

## Design notes

- **Bicep-first** (compile to ARM with `az bicep build` if you need JSON templates).
- **Existing accounts are never deleted**; deployment only adds the runtime environment and optional runbook/schedule.
- Graph **application** permissions cannot be granted with pure Automation ARM resources; use `Assign-GraphAppRoles.ps1`.
- Teams webhook URLs stay in the runbook invocation (not Automation variables), matching the current product design.
- If a package import fails (dependencies or gallery issues), re-import from the portal **Runtime environments** experience or pin versions via `graphModuleVersions`.

## Manual portal alternative

See [Setup Azure Automation Account](../Documentation/02-Setup-AzureAutomationAccount.md) and [Setup Azure Automation Runbook](../Documentation/03-Setup-AzureAutomation-Runbook.md).
