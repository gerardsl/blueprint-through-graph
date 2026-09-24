# Blueprint through Microsoft Graph

Create a Microsoft Entra **agent identity blueprint**, its **blueprint principal**, and an **agent identity** using Microsoft Graph.

[`New-AgentIdentityGraph.ps1`](New-AgentIdentityGraph.ps1) is a commented PowerShell example. It previews the requests by default and creates objects only when you pass `-Execute` and confirm.

## What it does

| Step | Microsoft Graph v1.0 request | Result |
| --- | --- | --- |
| 1 | `POST /applications/microsoft.graph.agentIdentityBlueprint` | Creates the blueprint application with an explicit sponsor. |
| 2 | `POST /servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal` | Creates the blueprint's principal in the tenant. |
| 3 | `POST /servicePrincipals/microsoft.graph.agentIdentity` | Creates an individual agent identity linked to the blueprint. |
| 4 | `GET /servicePrincipals/{id}/microsoft.graph.agentIdentity` | Reads back and verifies the identity. |

The blueprint's **application object ID** and **appId (client ID)** are different identifiers. Steps 2 and 3 use the returned **appId** to establish the relationship, not the application object ID.

## Requirements and permissions

- **PowerShell 7 or later.**
- **Microsoft.Graph.Authentication** for execution. Offline preview does not require the module or a sign-in.
- A work or school account in a tenant you are authorised to modify. This example supports the **Global** Microsoft Graph environment and **delegated** authentication.
- The target tenant ID and the object ID of an existing **sponsor user**. Supply the user's object ID, not their email address.
- **Agent ID Developer** or **Agent ID Administrator** for creating the blueprint and its principal.

The Graph client also needs administrator consent for these **delegated permissions**:

| Permission | Purpose |
| --- | --- |
| `AgentIdentityBlueprint.Create` | Create the blueprint application. |
| `AgentIdentityBlueprintPrincipal.Create` | Create its principal in the tenant. |
| `AgentIdentity.Create.All` | Create the agent identity. |
| `AgentIdentity.Read.All` | Read the identity back for verification. |

The signed-in creator becomes an owner of the new blueprint/principal, which supports creating its child identities with the required Graph permissions. An Entra role assignment does **not** replace consent to those permissions.

Do not use a connection containing `Directory.AccessAsUser.All`; the Agent ID APIs reject that permission. The script checks the connected tenant, authentication type, cloud environment, and all four scopes before sending creation requests.

## Run the example

### 1. Get the script and preview the requests

Clone this repository, or download the script and open PowerShell in its directory.

```powershell
git clone https://github.com/gerardsl/blueprint-through-graph.git
Set-Location .\blueprint-through-graph
```

Choose the target tenant, sponsor, and display names:

```powershell
$parameters = @{
    TenantId = [guid](Read-Host 'Target tenant ID')
    SponsorUserId = [guid](Read-Host 'Sponsor user object ID')
    BlueprintDisplayName = 'Document Assistant Blueprint'
    AgentDisplayName = 'Document Assistant'
}

.\New-AgentIdentityGraph.ps1 @parameters |
    ConvertTo-Json -Depth 8
```

This returns **`PLAN ONLY`** with the planned requests. It makes no Graph calls and creates nothing. The display names are examples; use your own names.

### 2. Connect to Microsoft Graph

Install the authentication module only if it is missing, then import it:

```powershell
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
}
Import-Module Microsoft.Graph.Authentication

$scopes = @(
    'AgentIdentityBlueprint.Create'
    'AgentIdentityBlueprintPrincipal.Create'
    'AgentIdentity.Create.All'
    'AgentIdentity.Read.All'
)

Connect-MgGraph -TenantId $parameters.TenantId -Scopes $scopes `
    -Environment Global -ContextScope Process -NoWelcome
```

Sign in with the intended account and complete any required administrator-approved consent beforehand. The script itself does not sign in, switch tenants, or grant itself permissions.

### 3. Execute once and review the result

> **This creates three real directory objects.** Review the target tenant, names, and sponsor before accepting the confirmation.

```powershell
$created = .\New-AgentIdentityGraph.ps1 @parameters -Execute
$created | Format-List
```

You can also use `-Execute -WhatIf` to preview without Graph calls:

```powershell
.\New-AgentIdentityGraph.ps1 @parameters -Execute -WhatIf
```

After a successful read-back, the script returns **`CREATED AND VERIFIED`**, together with:

- `BlueprintObjectId` and `BlueprintAppId`
- `BlueprintPrincipalObjectId`
- `AgentIdentityObjectId`, `AgentDisplayName`, and `ServicePrincipalType`
- `TenantId`, `SponsorUserId`, and a read-only `VerificationUri`

Verification checks that the returned identity has the expected object ID, display name, parent blueprint appId, and `ServiceIdentity` type.

Disconnect when finished:

```powershell
Disconnect-MgGraph
```

For parameter help:

```powershell
Get-Help .\New-AgentIdentityGraph.ps1 -Full
```

## Failures and cleanup

**Each execution creates new objects.** The script does not reuse objects by display name and is not an idempotent deployment tool.

If a request fails, the script stops and reports the stage and any object IDs already returned. A creation request may have succeeded even if its response or the later read-back failed. Inspect those IDs before retrying; do not rerun the whole workflow blindly.

There is no automatic rollback or cleanup. If the objects are no longer needed, have an authorised owner or administrator review their dependencies and remove only the objects created by this run, using their recorded IDs.

Follow your organisation's execution policy. Do not expose access tokens or private tenant details while recording or sharing output.

## Scope: identity provisioning only

The objective is to **create and verify the identity objects**. It is **not** to:

- Configure runtime credentials, certificates, secrets, or federation.
- Grant the agent access to APIs, files, mailboxes, or other resources.
- Deploy or run an agent application.
- Provision an agent user.
- Register or publish the agent in Agent 365 or a catalog.

The output flags `CredentialsCreated`, `ResourceGrantsCreated`, `RuntimeConfigured`, and `Agent365RegistrationVerified` are deliberately **`False`**. They describe work outside this script's scope, not failed identity creation.
