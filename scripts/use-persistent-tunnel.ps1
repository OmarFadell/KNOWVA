# Hosts a persistent dev tunnel instead of the ephemeral one, when asked to.
#
# WHY THIS SCRIPT EXISTS AT ALL
#
# VS Code tasks cannot read a .env file. `${env:FOO}` in tasks.json resolves
# against the environment VS Code itself was launched with, not against
# env/.env.local, and tasks.json has no conditionals -- `dependsOn` is a static
# list. So "run a different tunnel depending on a variable in .env.local" cannot
# be expressed in tasks.json alone. It needs a wrapper that reads the file and
# decides. This is that wrapper.
#
# WHAT IT DOES, AND THE ONE COMPROMISE IN IT
#
# It runs *after* the Toolkit's own "Start local tunnel" task, and corrects what
# that task wrote:
#
#   PERSISTENT_TUNNEL_HOST unset/empty -> does nothing at all, and the ephemeral
#                                         tunnel the Toolkit just created is what
#                                         gets used. Byte-for-byte the old
#                                         behaviour.
#   PERSISTENT_TUNNEL_HOST set         -> rewrites BOT_DOMAIN / BOT_ENDPOINT in
#                                         env/.env.local to the persistent host,
#                                         then hosts that tunnel in the
#                                         foreground so VS Code owns its
#                                         lifetime.
#
# THE COMPROMISE: in persistent mode the Toolkit has already created an
# ephemeral tunnel by the time this runs, and that tunnel is then ignored. It is
# wasted, and it is deliberate -- a shell task cannot invoke, replace or skip a
# `teamsfx`-type task, so the only way to keep the unset path *exactly* as it was
# is to let it run and then override it. If the waste matters more than a pristine
# fallback, remove "Start local tunnel" from the two dependsOn chains in
# .vscode/tasks.json; this script does not care either way, and the comments in
# tasks.json say the same.
#
# IT NEVER CREATES A TUNNEL. Only `devtunnel host`. Creating, naming, expiring
# and deleting the tunnel stay yours.

param(
  # Defaults to env/.env.local relative to the repo root (this script's parent).
  [string]$EnvFile
)

$ErrorActionPreference = "Stop"

if (-not $EnvFile) {
  $repoRoot = Split-Path -Parent $PSScriptRoot
  $EnvFile = Join-Path $repoRoot "env/.env.local"
}

# The line the tasks.json problemMatcher watches for. Printing it is what tells
# VS Code this task is "up" and the chain may continue to Provision. It is
# printed on BOTH paths -- a background task that never signals readiness leaves
# the whole F5 chain hanging.
$readySentinel = "PERSISTENT_TUNNEL_TASK_READY"

function Read-EnvFile([string]$path) {
  $map = @{}
  if (-not (Test-Path $path)) { return $map }

  foreach ($line in Get-Content -LiteralPath $path) {
    $trimmed = $line.Trim()
    if (-not $trimmed -or $trimmed.StartsWith("#")) { continue }

    $eq = $trimmed.IndexOf("=")
    if ($eq -lt 1) { continue }

    $key = $trimmed.Substring(0, $eq).Trim()
    $value = $trimmed.Substring($eq + 1).Trim().Trim('"').Trim("'")
    $map[$key] = $value
  }
  return $map
}

# Rewrites KEY=value in place, preserving every other line and the file's order.
# Appends the key if it is not already present. Writes UTF-8 with no BOM,
# because a BOM at the top of a .env file becomes part of the first key name and
# the Toolkit then silently fails to find it.
function Set-EnvValue([string]$path, [string]$key, [string]$value) {
  $lines = @()
  if (Test-Path $path) { $lines = @(Get-Content -LiteralPath $path) }

  $found = $false
  for ($i = 0; $i -lt $lines.Count; $i++) {
    $trimmed = $lines[$i].Trim()
    if ($trimmed.StartsWith("#")) { continue }

    $eq = $trimmed.IndexOf("=")
    if ($eq -lt 1) { continue }

    if ($trimmed.Substring(0, $eq).Trim() -eq $key) {
      $lines[$i] = "$key=$value"
      $found = $true
      break
    }
  }

  if (-not $found) { $lines += "$key=$value" }

  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($path, ($lines -join "`n") + "`n", $utf8NoBom)
}

# Derives the devtunnel tunnel id from its public hostname.
#
#   xf41ztqn-3978.uks1.devtunnels.ms  ->  xf41ztqn.uks1
#
# Dev tunnel hostnames are <tunnelId>-<port>.<cluster>.devtunnels.ms, and
# `devtunnel host` accepts "<tunnelId>.<cluster>". Deriving it means the host
# name is the single thing you configure, with nothing second to keep in sync.
# Set PERSISTENT_TUNNEL_ID if you would rather name the tunnel explicitly (e.g.
# "knowva_tunnel") -- that wins over anything derived here.
function Get-TunnelIdFromHost([string]$hostName) {
  $suffix = ".devtunnels.ms"
  if (-not $hostName.EndsWith($suffix)) { return $null }

  $stem = $hostName.Substring(0, $hostName.Length - $suffix.Length)   # xf41ztqn-3978.uks1
  $labels = $stem.Split(".")
  if ($labels.Count -lt 2) { return $null }

  $idWithPort = $labels[0]                                            # xf41ztqn-3978
  $cluster = $labels[1]                                               # uks1

  $tunnelId = [System.Text.RegularExpressions.Regex]::Replace($idWithPort, "-\d+$", "")
  if (-not $tunnelId) { return $null }

  return "$tunnelId.$cluster"
}

# --- Decide which mode we are in --------------------------------------------
#
# A NON-EMPTY process environment variable wins over the file, which is handy for
# pointing at a different tunnel for one session without editing anything:
#   $env:PERSISTENT_TUNNEL_HOST = "other-3978.uks1.devtunnels.ms"; code .
#
# It cannot be used the other way round. PowerShell deletes an environment
# variable when you assign "" to it, so there is no such thing as "set but
# empty" here and $env:PERSISTENT_TUNNEL_HOST = "" simply falls through to the
# file. To go back to ephemeral tunnels, clear the value in env/.env.local.
$envValues = Read-EnvFile $EnvFile

$persistentHost = $env:PERSISTENT_TUNNEL_HOST
if (-not $persistentHost) { $persistentHost = $envValues["PERSISTENT_TUNNEL_HOST"] }
if ($persistentHost) { $persistentHost = $persistentHost.Trim() }

if (-not $persistentHost) {
  Write-Host "PERSISTENT_TUNNEL_HOST is not set -- using the ephemeral dev tunnel the Toolkit just created." -ForegroundColor DarkGray
  Write-Host "  (Set it in env/.env.local to reuse a tunnel whose hostname does not change.)" -ForegroundColor DarkGray
  Write-Host $readySentinel
  exit 0
}

# --- Persistent mode ---------------------------------------------------------
$tunnelId = $env:PERSISTENT_TUNNEL_ID
if (-not $tunnelId) { $tunnelId = $envValues["PERSISTENT_TUNNEL_ID"] }
if ($tunnelId) { $tunnelId = $tunnelId.Trim() }

if (-not $tunnelId) {
  $tunnelId = Get-TunnelIdFromHost $persistentHost
}

if (-not $tunnelId) {
  Write-Host "PERSISTENT_TUNNEL_HOST is set to '$persistentHost', but no tunnel id could be worked out from it." -ForegroundColor Red
  Write-Host "Expected a hostname like xf41ztqn-3978.uks1.devtunnels.ms, or set PERSISTENT_TUNNEL_ID to the tunnel's name (e.g. knowva_tunnel)." -ForegroundColor Red
  exit 1
}

$devtunnel = Get-Command devtunnel -ErrorAction SilentlyContinue
if (-not $devtunnel) {
  Write-Host "'devtunnel' is not on PATH, so the persistent tunnel cannot be hosted." -ForegroundColor Red
  Write-Host "Install it (winget install Microsoft.devtunnel) and reopen VS Code so it picks up the new PATH," -ForegroundColor Red
  Write-Host "or clear PERSISTENT_TUNNEL_HOST in env/.env.local to go back to ephemeral tunnels." -ForegroundColor Red
  exit 1
}

Write-Host "Using persistent dev tunnel." -ForegroundColor Cyan
Write-Host "  host:      $persistentHost"
Write-Host "  tunnel id: $tunnelId"
Write-Host "  env file:  $EnvFile"

# These two are what everything downstream reads:
#   BOT_DOMAIN   -> infra/local.parameters.json -> local.bicep botAppDomain
#                   -> the Bot Framework messaging endpoint
#   BOT_ENDPOINT -> m365agents.local.yml -> .localConfigs -> config.ts
#                   -> the GitHub OAuth callback URL
# Correcting them here means both agree, in either tunnel mode, from one value.
Set-EnvValue $EnvFile "BOT_DOMAIN" $persistentHost
Set-EnvValue $EnvFile "BOT_ENDPOINT" "https://$persistentHost"

Write-Host "  BOT_DOMAIN   = $persistentHost" -ForegroundColor Green
Write-Host "  BOT_ENDPOINT = https://$persistentHost" -ForegroundColor Green
Write-Host "  GitHub OAuth callback = https://$persistentHost/auth/github/callback" -ForegroundColor Green

# Signalled before `devtunnel host` rather than after, because devtunnel's
# ready-message wording is not something to depend on. The small race that
# creates is harmless: the next tasks (Provision, Deploy) only write config,
# and nothing reaches the tunnel until "Start application" runs.
Write-Host $readySentinel

# Foreground on purpose. VS Code owns this process, so terminating the task (or
# stopping the debug session) stops hosting the tunnel -- no stray background
# process, and nothing to remember to kill by hand.
Write-Host "Hosting tunnel $tunnelId (Ctrl+C or terminating this task stops it)..." -ForegroundColor Cyan
& devtunnel host $tunnelId
exit $LASTEXITCODE
