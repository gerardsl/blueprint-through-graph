#requires -Version 7.0

<#
.SYNOPSIS
Creates a blueprint, its principal and one agent identity with Microsoft Graph.
.DESCRIPTION
Prints an offline plan by default. Execution uses an existing delegated
Microsoft Graph connection, verifies the tenant and scopes, and asks for
confirmation. It creates no credentials, resource grants or agent user.
Each execution creates new objects. Do not rerun after a partial failure;
inspect the reported object IDs first.
.NOTES
Permissions required for -Execute:

The Microsoft Graph client needs administrator consent for these delegated scopes:
  - AgentIdentityBlueprint.Create          Create the blueprint application.
  - AgentIdentityBlueprintPrincipal.Create Create its principal in the tenant.
  - AgentIdentity.Create.All               Create the agent identity.
  - AgentIdentity.Read.All                 Read the identity back for verification.

The signed-in user needs Agent ID Developer or Agent ID Administrator for
blueprint/principal creation. The creator becomes their owner, which allows
creation of child agent identities with the required Graph permissions.
An Entra role assignment does not replace consent to the Graph scopes.

Connect-MgGraph must already be connected to the specified tenant using a
delegated Global Graph session. Do not include Directory.AccessAsUser.All;
Agent ID APIs reject that permission.

The offline plan and -Execute -WhatIf need no Graph permissions or sign-in.
.PARAMETER TenantId
The directory in which to create the objects. The existing Graph connection
must match this tenant; the script never switches tenants.
.PARAMETER SponsorUserId
The object ID of an existing user who will sponsor both the blueprint and
the agent identity. This is a user object ID, not a name or email address.
.PARAMETER BlueprintDisplayName
A readable name for the shared identity blueprint.
.PARAMETER AgentDisplayName
A readable name for the individual agent identity created from the blueprint.
.PARAMETER Execute
Enables the creation workflow. Without this switch, only an offline plan is
returned. Execution also honours PowerShell's -Confirm and -WhatIf controls.
.EXAMPLE
.\New-AgentIdentityGraph.ps1 -TenantId $TenantId -SponsorUserId $SponsorUserId `
    -BlueprintDisplayName 'Document Assistant Blueprint' `
    -AgentDisplayName 'Document Assistant'
.EXAMPLE
.\New-AgentIdentityGraph.ps1 @parameters -Execute
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][guid]$TenantId,
    [Parameter(Mandatory)][guid]$SponsorUserId,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$BlueprintDisplayName,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$AgentDisplayName,
    [switch]$Execute
)

# Stop on undefined variables and cmdlet errors instead of continuing with
# missing data or treating a failed Graph request as a successful creation.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Validate the target identifiers and names before inspecting authentication
# or sending requests. Names are labels; GUIDs establish the relationships.
if ($TenantId -eq [guid]::Empty -or $SponsorUserId -eq [guid]::Empty) {
    throw 'TenantId and SponsorUserId must be explicit, non-empty GUIDs.'
}
foreach ($name in @($BlueprintDisplayName, $AgentDisplayName)) {
    if ([string]::IsNullOrWhiteSpace($name) -or $name.Length -gt 256 -or
        $name -ne $name.Trim() -or $name -match '[\x00-\x1F\x7F]') {
        throw 'Display names must be 1-256 characters without surrounding whitespace or control characters.'
    }
}

# Use the stable, typed Graph endpoints and make the OData version explicit.
$graph = 'https://graph.microsoft.com/v1.0'
$headers = @{ 'OData-Version' = '4.0' }
$requiredScopes = @(
    # Create the shared blueprint application.
    'AgentIdentityBlueprint.Create'
    # Create the blueprint's service-principal representation in this tenant.
    'AgentIdentityBlueprintPrincipal.Create'
    # Create the individual agent identity under the blueprint.
    'AgentIdentity.Create.All'
    # Read the new agent identity back to verify its properties.
    'AgentIdentity.Read.All'
)
# @odata.bind links the new objects to an existing sponsor user. It does not
# create a user or grant the agent access to that user's resources.
$sponsor = "$graph/users/$SponsorUserId"
$blueprintBody = @{
    displayName = $BlueprintDisplayName
    'sponsors@odata.bind' = @($sponsor)
}
# Describe the three creation requests and the verification request without
# executing them. Placeholder IDs will be replaced by actual Graph responses.
$plan = [pscustomobject]@{
    Mode = 'PLAN ONLY'
    TenantId = $TenantId.ToString()
    SponsorUserId = $SponsorUserId.ToString()
    RequiredDelegatedScopes = $requiredScopes
    Requests = @(
        [pscustomobject]@{
            Method = 'POST'
            Uri = "$graph/applications/microsoft.graph.agentIdentityBlueprint"
            Body = $blueprintBody
        }
        [pscustomobject]@{
            Method = 'POST'
            Uri = "$graph/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal"
            Body = @{ appId = '<created-blueprint-appId>' }
        }
        [pscustomobject]@{
            Method = 'POST'
            Uri = "$graph/servicePrincipals/microsoft.graph.agentIdentity"
            Body = @{
                displayName = $AgentDisplayName
                agentIdentityBlueprintId = '<created-blueprint-appId>'
                'sponsors@odata.bind' = @($sponsor)
            }
        }
        [pscustomobject]@{
            Method = 'GET'
            Uri = "$graph/servicePrincipals/<created-agent-object-id>/microsoft.graph.agentIdentity"
        }
    )
    TenantActionsPerformed = $false
    Notice = 'Creates new directory objects only. No credentials, resource grants, runtime or Agent 365 registration are configured.'
}
# The default path exits here, so previewing requires neither the Graph module
# nor a sign-in. WhatIf or declined confirmation also returns without requests.
if (-not $Execute) { return $plan }
if (-not $PSCmdlet.ShouldProcess(
    "$AgentDisplayName under $BlueprintDisplayName in tenant $TenantId",
    'Create a blueprint, blueprint principal and agent identity, then verify the identity'
)) {
    return $plan
}

# Reuse a connection the operator established explicitly with Connect-MgGraph.
# Do not install modules, sign in, or request additional consent automatically.
foreach ($command in @('Get-MgContext', 'Invoke-MgGraphRequest')) {
    if (-not (Get-Command $command -ErrorAction SilentlyContinue)) {
        throw 'Import Microsoft.Graph.Authentication and connect with the documented delegated scopes before using -Execute.'
    }
}
$context = Get-MgContext
if ($null -eq $context) {
    throw 'No Microsoft Graph connection. Run Connect-MgGraph with the intended TenantId and delegated scopes first.'
}
# Ensure a valid session cannot accidentally target a different directory or
# cloud. This example uses a signed-in user's delegated permissions.
$connectedTenant = [guid]::Empty
if (-not [guid]::TryParse([string]$context.TenantId, [ref]$connectedTenant) -or
    $connectedTenant -ne $TenantId) {
    throw 'The Microsoft Graph connection is for a different tenant. No creation request was sent.'
}
if ($context.AuthType -ne 'Delegated' -or $context.Environment -ne 'Global') {
    throw 'This example requires a delegated connection to the Microsoft Graph Global environment.'
}
# Agent ID rejects this broad legacy scope. Check the required scopes as well;
# Graph still enforces the caller's roles/ownership when processing each request.
if (@($context.Scopes) -contains 'Directory.AccessAsUser.All') {
    throw 'Reconnect without Directory.AccessAsUser.All; the Agent ID APIs reject that delegated permission.'
}
$missingScopes = @($requiredScopes | Where-Object { $_ -notin @($context.Scopes) })
if ($missingScopes.Count) {
    throw "Reconnect with the required delegated scopes: $($missingScopes -join ', '). No creation request was sent."
}

# Accept only an object response with a usable ID, and check its OData type
# when supplied. A malformed response must not feed the next creation step.
function Assert-GraphObject {
    param([object]$Value, [string]$ExpectedType)
    if ($Value -isnot [System.Collections.IDictionary]) {
        throw "Graph did not return an object for $ExpectedType."
    }
    $id = [guid]::Empty
    if (-not $Value.Contains('id') -or
        -not [guid]::TryParse([string]$Value['id'], [ref]$id) -or $id -eq [guid]::Empty) {
        throw "Graph did not return a valid object ID for $ExpectedType."
    }
    if ($Value.Contains('@odata.type') -and $Value['@odata.type'] -ine "#microsoft.graph.$ExpectedType") {
        throw "Graph returned a different object type instead of $ExpectedType."
    }
    return $id.ToString()
}

# Compare identifiers as GUIDs rather than names or case-sensitive strings,
# so the returned objects must belong to the blueprint/identity we requested.
function Assert-MatchingGuid {
    param([object]$Value, [guid]$Expected, [string]$Property)
    $actual = [guid]::Empty
    if (-not [guid]::TryParse([string]$Value, [ref]$actual) -or $actual -ne $Expected) {
        throw "Graph returned an unexpected $Property."
    }
}

# Keep the returned, non-secret IDs and the current stage for recovery guidance
# if a later operation fails. These requests are not one atomic transaction.
$created = [ordered]@{}
$stage = 'creating the blueprint'
try {
    # 1. Create the blueprint application with its sponsor. The Graph SDK uses
    # the current connection for authentication; no bearer token is handled here.
    $blueprint = Invoke-MgGraphRequest -Method POST `
        -Uri "$graph/applications/microsoft.graph.agentIdentityBlueprint" `
        -Headers $headers -ContentType 'application/json' `
        -Body ($blueprintBody | ConvertTo-Json -Depth 5) -OutputType Hashtable -ErrorAction Stop
    $created['BlueprintObjectId'] = Assert-GraphObject $blueprint 'agentIdentityBlueprint'
    # Keep both IDs: id identifies the directory object; appId is the client ID
    # used to link the blueprint principal and child agent identity.
    $appId = [guid]::Empty
    if (-not $blueprint.Contains('appId') -or
        -not [guid]::TryParse([string]$blueprint['appId'], [ref]$appId) -or $appId -eq [guid]::Empty) {
        throw 'Graph did not return a valid blueprint appId.'
    }
    $created['BlueprintAppId'] = $appId.ToString()
    Write-Host "Blueprint created: object ID $($created['BlueprintObjectId']); appId $appId"

    # 2. Create the blueprint principal using appId, not BlueprintObjectId.
    # This is the blueprint's presence in the tenant, not the agent itself.
    $stage = 'creating the blueprint principal'
    $principal = Invoke-MgGraphRequest -Method POST `
        -Uri "$graph/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal" `
        -Headers $headers -ContentType 'application/json' `
        -Body (@{ appId = $appId.ToString() } | ConvertTo-Json) -OutputType Hashtable -ErrorAction Stop
    $created['BlueprintPrincipalObjectId'] = Assert-GraphObject $principal 'agentIdentityBlueprintPrincipal'
    Assert-MatchingGuid $principal['appId'] $appId 'blueprint principal appId'
    Write-Host "Blueprint principal created: $($created['BlueprintPrincipalObjectId'])"

    # 3. Create a purpose-built agent identity. The typed endpoint distinguishes
    # it from a conventional service principal; the parent link uses appId again.
    $stage = 'creating the agent identity'
    $agentBody = @{
        displayName = $AgentDisplayName
        agentIdentityBlueprintId = $appId.ToString()
        'sponsors@odata.bind' = @($sponsor)
    }
    $agent = Invoke-MgGraphRequest -Method POST `
        -Uri "$graph/servicePrincipals/microsoft.graph.agentIdentity" `
        -Headers $headers -ContentType 'application/json' `
        -Body ($agentBody | ConvertTo-Json -Depth 5) -OutputType Hashtable -ErrorAction Stop
    $created['AgentIdentityObjectId'] = Assert-GraphObject $agent 'agentIdentity'
    Assert-MatchingGuid $agent['agentIdentityBlueprintId'] $appId 'agentIdentityBlueprintId'
    Write-Host "Agent identity created: $($created['AgentIdentityObjectId'])"

    # 4. Read the identity through its typed endpoint, independently of the POST
    # response. Verify its ID, parent, requested name and ServiceIdentity type.
    $stage = 'reading back the agent identity'
    $verificationUri = "$graph/servicePrincipals/$($created['AgentIdentityObjectId'])/microsoft.graph.agentIdentity"
    $verified = Invoke-MgGraphRequest -Method GET -Uri $verificationUri `
        -Headers $headers -OutputType Hashtable -ErrorAction Stop
    $verifiedId = Assert-GraphObject $verified 'agentIdentity'
    Assert-MatchingGuid $verifiedId ([guid]$created['AgentIdentityObjectId']) 'agent identity object ID'
    Assert-MatchingGuid $verified['agentIdentityBlueprintId'] $appId 'agentIdentityBlueprintId'
    if ($verified['displayName'] -cne $AgentDisplayName -or
        $verified['servicePrincipalType'] -ne 'ServiceIdentity') {
        throw 'The identity read-back did not match the requested name and ServiceIdentity type.'
    }
} catch {
    # Surface the failure with any IDs already returned. Even a lost response
    # can follow successful creation, so never assume a blind retry is safe.
    $knownIds = ($created.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
    if (-not $knownIds) { $knownIds = 'No object IDs were returned.' }
    throw [InvalidOperationException]::new(
        "Graph workflow stopped while $stage. A creation request may have succeeded. " +
        "Inspect the tenant before retrying; no rollback or automatic cleanup was attempted. " +
        "Known IDs: $knownIds. Graph detail: $($_.Exception.Message)", $_.Exception
    )
}

# Return a concise result only after read-back succeeds. The false flags make
# clear that directory provisioning is not runtime setup or resource consent.
[pscustomobject]@{
    Mode = 'CREATED AND VERIFIED'
    TenantId = $TenantId.ToString()
    BlueprintObjectId = $created['BlueprintObjectId']
    BlueprintAppId = $created['BlueprintAppId']
    BlueprintPrincipalObjectId = $created['BlueprintPrincipalObjectId']
    AgentIdentityObjectId = $created['AgentIdentityObjectId']
    AgentDisplayName = $verified['displayName']
    ServicePrincipalType = $verified['servicePrincipalType']
    SponsorUserId = $SponsorUserId.ToString()
    VerificationUri = $verificationUri
    CredentialsCreated = $false
    ResourceGrantsCreated = $false
    RuntimeConfigured = $false
    Agent365RegistrationVerified = $false
}

# Objective: create and verify the blueprint, its principal and an agent identity.
# This script is not intended to configure runtime credentials, grant resource
# access, or deploy an agent application. Those are separate follow-up tasks.
# The false flags above describe these intentional scope limits, not failed
# identity creation. Agent 365 registration is also not performed or verified.
