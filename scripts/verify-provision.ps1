# Verifies that provisioning actually produced what it claims.
# "Provision succeeded" is not evidence -- this checks the artifacts themselves.

param(
  [Parameter(Mandatory = $true)][string]$BotId,
  [Parameter(Mandatory = $true)][string]$ResourceGroup,
  [string]$ConnectionName = "graph"
)

$ErrorActionPreference = "Stop"
$failures = @()

Write-Host "Verifying provision artifacts..." -ForegroundColor Cyan

# 1. Entra app exists
$app = az ad app show --id $BotId --query "displayName" -o tsv 2>$null
if (-not $app) {
  $failures += "Entra app $BotId not found. aadApp/create did not complete."
} else {
  Write-Host "  [ok] Entra app: $app" -ForegroundColor Green
}

# 1b. SSO manifest (aad.manifest.json) was actually applied by aadApp/update.
# Half-applied SSO config (missing identifierUris or the access_as_user scope)
# fails silently at runtime as a token-exchange error, not at provision --
# so check the artifacts here instead.
if ($app) {
  $expectedIdentifierUri = "api://botid-$BotId"
  $identifierUris = az ad app show --id $BotId --query "identifierUris" -o json 2>$null | ConvertFrom-Json
  if (-not $identifierUris -or ($identifierUris -notcontains $expectedIdentifierUri)) {
    $failures += "Entra app $BotId is missing identifierUris entry '$expectedIdentifierUri'. aadApp/update (aad.manifest.json) did not apply."
  } else {
    Write-Host "  [ok] identifierUris: $expectedIdentifierUri" -ForegroundColor Green
  }

  $accessAsUserScope = az ad app show --id $BotId --query "api.oauth2PermissionScopes[?value=='access_as_user'] | [0]" -o json 2>$null | ConvertFrom-Json
  if (-not $accessAsUserScope) {
    $failures += "Entra app $BotId has no 'access_as_user' OAuth2 permission scope. aadApp/update (aad.manifest.json) did not apply, or oauth2Permissions used the wrong manifest schema."
  } elseif (-not $accessAsUserScope.isEnabled) {
    $failures += "Entra app $BotId has an 'access_as_user' scope but it is disabled."
  } else {
    Write-Host "  [ok] OAuth2 permission scope: access_as_user (enabled, id=$($accessAsUserScope.id))" -ForegroundColor Green
  }

  # preAuthorizedApplications is what lets Teams clients request a token for
  # this scope without a per-user consent prompt. An empty list here means
  # Teams has no scope to request a token against -- token exchange fails
  # silently at runtime (this is exactly the bug this check exists to catch).
  $preAuthApps = az ad app show --id $BotId --query "api.preAuthorizedApplications[].appId" -o json 2>$null | ConvertFrom-Json
  if (-not $preAuthApps -or @($preAuthApps).Count -eq 0) {
    $failures += "Entra app $BotId has no preAuthorizedApplications. Teams has no scope to request a token against -- token exchange will fail silently. aadApp/update (aad.manifest.json) did not apply, or preAuthorizedApplications used the wrong manifest schema."
  } else {
    Write-Host "  [ok] preAuthorizedApplications: $(@($preAuthApps).Count) app(s) -- $($preAuthApps -join ', ')" -ForegroundColor Green

    # Sanity-check that the scope each pre-authorized app is granted actually
    # matches our access_as_user scope id, not some stale/unrelated GUID.
    if ($accessAsUserScope) {
      $wrongScopeApps = az ad app show --id $BotId --query "api.preAuthorizedApplications[?!contains(delegatedPermissionIds, '$($accessAsUserScope.id)')].appId" -o json 2>$null | ConvertFrom-Json
      if ($wrongScopeApps -and @($wrongScopeApps).Count -gt 0) {
        $failures += "preAuthorizedApplications entries for $($wrongScopeApps -join ', ') do not reference the current access_as_user scope id ($($accessAsUserScope.id)). Stale scope id -- check AAD_APP_ACCESS_AS_USER_PERMISSION_ID."
      }
    }
  }
}

# 2. A bot resource is actually backed by this app id
$bots = az resource list --resource-group $ResourceGroup --resource-type Microsoft.BotService/botServices --query "[].name" -o tsv
$botName = $null
foreach ($b in $bots) {
  $msaAppId = az bot show --name $b --resource-group $ResourceGroup --query "properties.msaAppId" -o tsv 2>$null
  if ($msaAppId -eq $BotId) { $botName = $b; break }
}
if (-not $botName) {
  $failures += "No bot resource in $ResourceGroup is backed by app id $BotId. The ARM deployment did not create it."
} else {
  Write-Host "  [ok] Bot resource: $botName" -ForegroundColor Green

  # 3. Teams channel enabled
  $teams = az bot msteams show --name $botName --resource-group $ResourceGroup --query "name" -o tsv 2>$null
  if (-not $teams) {
    $failures += "Teams channel not enabled on $botName. Teams will report 'Invalid Bot'."
  } else {
    Write-Host "  [ok] Teams channel enabled" -ForegroundColor Green
  }

  # 4. OAuth connection present (needed for SSO milestone)
  $conn = az bot authsetting show --name $botName --resource-group $ResourceGroup --setting-name $ConnectionName --query "name" -o tsv 2>$null
  if (-not $conn) {
    Write-Host "  [warn] OAuth connection '$ConnectionName' missing -- fine until the SSO milestone" -ForegroundColor Yellow
  } else {
    Write-Host "  [ok] OAuth connection: $ConnectionName" -ForegroundColor Green
  }

  # 5. Endpoint matches what we think it is
  $endpoint = az bot show --name $botName --resource-group $ResourceGroup --query "properties.endpoint" -o tsv
  Write-Host "  [info] Messaging endpoint: $endpoint" -ForegroundColor Gray
}

if ($failures.Count -gt 0) {
  Write-Host ""
  Write-Host "PROVISION VERIFICATION FAILED:" -ForegroundColor Red
  $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
  exit 1
}

Write-Host "All provision artifacts verified." -ForegroundColor Green
exit 0
