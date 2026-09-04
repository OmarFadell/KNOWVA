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
