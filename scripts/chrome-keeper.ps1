# Chrome keeper - run the operator's Chrome as a supervised background process
#
# Purpose: the personal mint harvests a token from a live chat tab, so a Chrome
# with --remote-debugging-port=9222 must be running on the capture profile. This
# script keeps one there: it starts Chrome, restarts it if it dies, and reopens
# the chat tab if the tab is closed or the profile lost its session.
#
# It does NOT replace `pnpm m365-personal login`: signing in, and clearing
# Microsoft's human-verification dialog, still need a visible window. Run this
# only after a login has succeeded and the profile holds a live session.
#
# IMPORTANT - headless does not work (measured 2026-10-23):
#   A headless=new Chrome starts, loads the chat page and still shows the
#   signed-in session, but the composer will not accept a click and the Send
#   button never appears - the page reports a verification/security check. The
#   gate is on interaction, not on page load, so a mint against it fails with
#   mint_not_rotated. Keep the window. See docs/personal-mint-design.md.
#
# Usage:
#   pwsh scripts/chrome-keeper.ps1                  # supervise a headed Chrome
#   pwsh scripts/chrome-keeper.ps1 -Once            # start once, no supervision
#   pwsh scripts/chrome-keeper.ps1 -Headed:$false   # headless (mints will fail)
#
# One Chrome per account (multi-account design 7.1): the profile and the
# debugging port belong to the account, so run one keeper per account:
#   pwsh scripts/chrome-keeper.ps1 -Account work
# -Account reads that account's cdpPort and profileDir from
# "$HOME\.config\opencode-m365\accounts.json" (or M365_ACCOUNTS_FILE); -Port and
# -ProfileDir still override them. Two Chromes on one profile, or two accounts
# on one port, are the two ways to break the mint. Without -Account the original
# single-account defaults (personal-profile, 9222) apply unchanged.
#
# Stop it with Ctrl-C; Ctrl-C does not close the Chrome it started.

[CmdletBinding()]
param(
  [string] $ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe",
  [int]    $Port = 9222,
  [string] $ProfileDir = "$HOME\.config\opencode-m365\personal-profile",
  [string] $ChatUrl = "https://m365.cloud.microsoft/chat",
  # Which account's Chrome to keep: its port and profile come from the registry.
  [string] $Account = "",
  # Headed is the default because headless cannot mint: the page loads and shows
  # the session, then refuses every click (see the header).
  [switch] $Headed = $true,
  [switch] $Once,
  [int]    $CheckIntervalSec = 20
)

$ErrorActionPreference = "Stop"

if ($Account -ne "") {
  $registryFile = if ($env:M365_ACCOUNTS_FILE) { $env:M365_ACCOUNTS_FILE } else { "$HOME\.config\opencode-m365\accounts.json" }
  if (-not (Test-Path $registryFile)) {
    throw "no account registry at $registryFile - add the account first (the proxy's /accounts/ui, or POST /accounts)"
  }
  $registry = Get-Content -Raw $registryFile | ConvertFrom-Json
  $entry = @($registry.accounts | Where-Object { $_.id -eq $Account }) | Select-Object -First 1
  if (-not $entry) {
    $known = (@($registry.accounts | ForEach-Object { $_.id }) -join ", ")
    throw "account '$Account' is not in accounts.json; known: $known"
  }
  if (-not $PSBoundParameters.ContainsKey("Port")) { $Port = [int] $entry.cdpPort }
  if (-not $PSBoundParameters.ContainsKey("ProfileDir")) { $ProfileDir = [string] $entry.profileDir }
  Write-Host "account '$Account': port $Port, profile $ProfileDir"
}

if (-not $Headed) {
  Write-Warning "Headless mode: a mint will fail (mint_not_rotated) - the page loads but refuses interaction. Use a headed Chrome."
}
if (-not $PSBoundParameters.ContainsKey("Headed")) { $Headed = $true }

function Write-Log([string] $msg) {
  Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg)
}

function Test-Cdp {
  try {
    $null = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/json/version" -TimeoutSec 3
    return $true
  } catch {
    return $false
  }
}

# A chat tab must exist or the harvest answers `no_chat_tab`. A tab sitting on a
# sign-in page means the profile's session is gone and only a human can fix it.
function Get-ChatTabState {
  try {
    $tabs = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/json" -TimeoutSec 3
  } catch {
    return "cdp_down"
  }
  $chat = @($tabs | Where-Object { $_.type -eq "page" -and $_.url -like "*m365.cloud.microsoft*" })
  if ($chat.Count -eq 0) { return "no_tab" }
  if (@($chat | Where-Object { $_.url -like "*login.microsoftonline.com*" -or $_.url -like "*/login*" }).Count -eq $chat.Count) {
    return "signed_out"
  }
  return "ok"
}

function Start-KeeperChrome {
  $args = @(
    "--remote-debugging-port=$Port",
    "--user-data-dir=$ProfileDir",
    "--no-first-run",
    "--no-default-browser-check"
  )
  if (-not $Headed) {
    # A headless Chrome is still a real Chrome: same profile, same cookies, same
    # session. It is NOT the Playwright launch that Cloudflare fingerprints.
    $args += "--headless=new"
  }
  $args += $ChatUrl
  Write-Log ("starting Chrome ({0}) on port {1}" -f ($(if ($Headed) { "headed" } else { "headless" })), $Port)
  Start-Process -FilePath $ChromePath -ArgumentList $args | Out-Null
}

function Open-ChatTab {
  Write-Log "no chat tab - opening one"
  try {
    Invoke-RestMethod -Method Put -Uri "http://127.0.0.1:$Port/json/new?$ChatUrl" -TimeoutSec 5 | Out-Null
  } catch {
    # /json/new is disabled in newer Chrome builds; restart instead.
    Write-Log "could not open a tab via CDP ($($_.Exception.Message)) - restarting Chrome"
    Get-Process chrome -ErrorAction SilentlyContinue |
      Where-Object { $_.CommandLine -like "*--user-data-dir=$ProfileDir*" } |
      Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-KeeperChrome
  }
}

if (-not (Test-Path $ChromePath)) { throw "chrome not found at $ChromePath - pass -ChromePath" }

if (Test-Cdp) {
  Write-Log "something already listens on $Port - attaching, not starting another Chrome"
} else {
  Start-KeeperChrome
  if ($Once) { Write-Log "started (once); exiting"; return }
  Start-Sleep -Seconds 5
}

if ($Once) { return }

Write-Log "supervising every $($CheckIntervalSec)s - Ctrl-C to stop"
$signedOutWarned = $false
while ($true) {
  Start-Sleep -Seconds $CheckIntervalSec
  $state = Get-ChatTabState
  switch ($state) {
    "cdp_down" {
      Write-Log "CDP is gone - restarting Chrome"
      Start-KeeperChrome
      $signedOutWarned = $false
      Start-Sleep -Seconds 5
    }
    "no_tab" { Open-ChatTab }
    "signed_out" {
      if (-not $signedOutWarned) {
        Write-Log "the chat tab is on a sign-in page: the profile's session is gone."
        Write-Log "a human has to sign in - stop this, run: pnpm m365-personal login, then start it again."
        $signedOutWarned = $true
      }
    }
    "ok" { $signedOutWarned = $false }
  }
}
