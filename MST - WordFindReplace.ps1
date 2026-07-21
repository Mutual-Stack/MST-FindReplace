<#
.SYNOPSIS
    Deletes YOUR OWN Microsoft Teams chat messages that match any pattern
    in $DeletePatterns. No reposts, no edits - matched messages are
    soft-deleted and logged. Built for cleanup of profanity (test phase)
    and future removal of leaked licenses/passwords/keys.

.IMPORTANT
    - Graph cannot edit a sent message in place. Deletion is the only
      clean server-side remediation.
    - Soft delete is surface-only. Compliance/eDiscovery copies remain.
      A leaked credential is STILL BURNED - rotate it regardless.
    - Only messages YOU sent are touched (delegated auth + sender check).
    - Recipients see the standard "This message was deleted" placeholder.
    - The audit CSV contains the original message text unless -MaskLog is
      used. For credential-removal runs, ALWAYS use -MaskLog and store the
      CSV somewhere access-controlled.

.REQUIREMENTS
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
    Run from a REGULAR PowerShell console, not ISE (WAM auth popup).

.EXAMPLE
    .\Remove-MyTeamsChatMessages.ps1 -DryRun                          # always first
    .\Remove-MyTeamsChatMessages.ps1 -DryRun -Since (Get-Date).AddMonths(-6)
    .\Remove-MyTeamsChatMessages.ps1                                  # live run
    .\Remove-MyTeamsChatMessages.ps1 -MaskLog                         # live, secrets masked in CSV
#>

[CmdletBinding()]
param(
    [switch]$DryRun,
    $Since,   # accepts datetime; untyped so a null default can't break closures/binding
    [string]$LogPath = ".\TeamsChatDelete_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

if ($Since) { $Since = [datetime]$Since }

# ============================================================================
#  DELETE PATTERNS - message is removed if it matches ANY of these.
#  Case-insensitive regex. Most-specific first is good hygiene but not
#  required (any single hit deletes).
# ============================================================================
$DeletePatterns = @(

    # --- TEST PHASE: professional cursing (wildcard-heavy on purpose) ---
    '\b\w*(?:f+u+c+k+|fuk|fuq|fcuk|fvck|phuck|fukc)\w*\b'
    '\b\w*(?:s+h+[i1y]+t+|hsit)\w*\b'
    '\b\w*(?:retard|retart)\w*\b'
    '\b\w*(?:pussy|pussie)\w*\b'
    # NOTE: deliberately NOT including the bare dick/dik wildcard pattern here.
    # \w* on both sides matches Dickson/Dickinson/Benedict etc. If you want it,
    # use the anchored form below instead:
    # '\b(?:dick|dik)(?:s|head|wad)?\b'

    # --- FUTURE PHASE: credentials/licenses (commented until needed) ---
    # '(?i)\b(?:password|passwd|pwd)\s*[:=]\s*\S+'          # password: xxxx
    # '(?i)\bapi[_-]?key\s*[:=]\s*\S+'                      # api_key=xxxx
    # '\b[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}(?:-[A-Z0-9]{5})?\b'  # XXXXX-XXXXX license keys
    # '\bAKIA[0-9A-Z]{16}\b'                                # AWS access key ID
    # '(?i)\bbearer\s+[a-z0-9\-\._~\+\/]{20,}=*'            # bearer tokens
    # 'eyJ[A-Za-z0-9_-]{10,}\.eyJ[A-Za-z0-9_-]{10,}'        # JWTs
)

# ============================================================================

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

Write-Host 'Connecting to Microsoft Graph (interactive sign-in as yourself)...' -ForegroundColor Cyan
Connect-MgGraph -Scopes 'Chat.ReadWrite','User.Read' -NoWelcome

# Hard guard: never reach the confirmation prompt unauthenticated
$ctx = Get-MgContext
if (-not $ctx -or -not $ctx.Account) { throw 'Graph auth failed - aborting before any changes.' }

$me   = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me'
$myId = $me.id
Write-Host "Authenticated as: $($me.displayName) <$($me.userPrincipalName)>" -ForegroundColor Green
Write-Host "Mode: $(if ($DryRun) {'DRY RUN - nothing changes'} else {'LIVE - matched messages WILL be deleted'})" -ForegroundColor $(if ($DryRun) {'Yellow'} else {'Red'})

if (-not $DryRun) {
    $confirm = Read-Host 'Type DELETE to confirm the live run'
    if ($confirm -cne 'DELETE') { Write-Host 'Aborted.'; return }
}

$stats     = @{ Chats=0; Scanned=0; Deleted=0; Skipped=0; Errors=0 }
$logHeader = $false

function Write-LogEntry {
    param($Entry)
    # Per-entry append: a Ctrl+C can never lose history
    $Entry | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -Append
}

function Invoke-GraphWithRetry {
    param([string]$Method,[string]$Uri,$Body)
    $attempt = 0
    while ($true) {
        try {
            if ($Body) {
                return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body ($Body | ConvertTo-Json -Depth 5) -ContentType 'application/json'
            } else {
                return Invoke-MgGraphRequest -Method $Method -Uri $Uri
            }
        } catch {
            $attempt++
            $resp   = $_.Exception.Response
            $status = if ($resp) { [int]$resp.StatusCode } else { 0 }
            if ($status -eq 429 -and $attempt -le 6) {
                $wait = 10; try { $wait = [int]($resp.Headers.GetValues('Retry-After') | Select-Object -First 1) } catch {}
                Write-Host "  Throttled. Waiting $wait s..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds $wait
            } elseif ($status -in 500,502,503,504 -and $attempt -le 4) {
                Start-Sleep -Seconds (5*$attempt)
            } else { throw }
        }
    }
}

function Get-PlainText {
    param([string]$Html)
    if ([string]::IsNullOrEmpty($Html)) { return '' }
    $t = $Html -replace '<[^>]+>',' '
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    return ($t -replace '\s+',' ').Trim()
}

function Get-LoggableText {
    param([string]$Plain, [string]$Pattern)
    # Always mask the matched text - console and CSV never show the hit itself
    return [regex]::Replace($Plain, $Pattern, '****', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

# --- enumerate chats ---
Write-Host 'Enumerating chats...' -ForegroundColor Cyan
$chats = [System.Collections.Generic.List[object]]::new()
$uri = 'https://graph.microsoft.com/v1.0/me/chats?$top=50&$expand=members'
while ($uri) { $p = Invoke-GraphWithRetry GET $uri; $p.value | ForEach-Object { $chats.Add($_) }; $uri = $p.'@odata.nextLink' }
Write-Host "Found $($chats.Count) chats.`n"

foreach ($chat in $chats) {
    $stats.Chats++
    $label = $chat.topic
    if ([string]::IsNullOrEmpty($label)) {
        $others = @($chat.members | Where-Object { $_.userId -ne $myId } | ForEach-Object { $_.displayName }) -join ', '
        $label  = if ($others) { $others } else { '(self/unknown)' }
    }
    Write-Host "[$($stats.Chats)/$($chats.Count)] $($chat.chatType): $label" -ForegroundColor Cyan

    $msgUri = "https://graph.microsoft.com/v1.0/me/chats/$($chat.id)/messages?`$top=50"
    $stop = $false
    while ($msgUri -and -not $stop) {
        $page = Invoke-GraphWithRetry GET $msgUri
        foreach ($msg in $page.value) {
            if ($Since -and $msg.createdDateTime -and [datetime]$msg.createdDateTime -lt $Since) { $stop = $true; break }
            if ($msg.messageType -ne 'message' -or $msg.deletedDateTime) { continue }
            if ($msg.from.user.id -ne $myId) { continue }   # only MY messages
            $stats.Scanned++

            $plain = Get-PlainText $msg.body.content

            $hit = $DeletePatterns | Where-Object { $plain -match $_ } | Select-Object -First 1
            if (-not $hit) { $stats.Skipped++; continue }

            $logText = Get-LoggableText -Plain $plain -Pattern $hit
            $entry = [pscustomobject]@{
                Timestamp = $msg.createdDateTime
                Chat      = $label
                Action    = ''
                MatchedOn = $hit
                Message   = $logText
                ChatId    = $chat.id
                MessageId = $msg.id
            }

            if ($DryRun) {
                $entry.Action = 'WOULD DELETE'
                Write-Host "  DELETE  <= [$hit] $($logText.Substring(0,[Math]::Min(80,$logText.Length)))" -ForegroundColor Yellow
            } else {
                try {
                    Invoke-GraphWithRetry POST "https://graph.microsoft.com/v1.0/me/chats/$($chat.id)/messages/$($msg.id)/softDelete" | Out-Null
                    $entry.Action = 'DELETED'; $stats.Deleted++
                    Write-Host "  DELETED <= [$hit]" -ForegroundColor Red
                    Start-Sleep -Milliseconds 300
                } catch {
                    $entry.Action = "ERROR: $($_.Exception.Message)"; $stats.Errors++
                }
            }
            Write-LogEntry $entry
        }
        $msgUri = $page.'@odata.nextLink'
    }
}

Write-Host "`n===== SUMMARY =====" -ForegroundColor Cyan
Write-Host "Chats:    $($stats.Chats)"
Write-Host "Scanned:  $($stats.Scanned)  (my messages)"
Write-Host "Deleted:  $($stats.Deleted)"
Write-Host "Skipped:  $($stats.Skipped)"
Write-Host "Errors:   $($stats.Errors)"
if (Test-Path $LogPath) { Write-Host "Audit log: $LogPath" -ForegroundColor Green }
if ($DryRun) { Write-Host "`nDry run complete. Review the CSV, then re-run without -DryRun." -ForegroundColor Yellow }
Disconnect-MgGraph | Out-Null
