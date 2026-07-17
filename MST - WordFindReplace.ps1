<#
.SYNOPSIS
    Cleans YOUR OWN Microsoft Teams chat messages two ways:
      1. REDACT  - swap matched words (e.g. profanity) for softer text
      2. DELETE  - remove the whole message if it hits a delete term

.IMPORTANT
    Graph CANNOT edit the body of a message you already sent. "Redact"
    is therefore implemented as: soft-delete the original, then POST a
    NEW message from you containing the cleaned text. Recipients see the
    "deleted" placeholder AND the new message. The repost:
      - gets a NEW timestamp (shows as sent now)
      - lands at the bottom of the chat, not the original position
      - loses reactions/replies that were on the original
    Soft delete is surface-only; compliance copies remain in the substrate.

    Only messages YOU sent are ever touched (delegated auth + sender check).

.REQUIREMENTS
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser

.EXAMPLE
    .\Redact-MyTeamsChatMessages.ps1 -DryRun        # ALWAYS run this first
    .\Redact-MyTeamsChatMessages.ps1                # live (prompts for DELETE)
#>

[CmdletBinding()]
param(
    [switch]$DryRun,
    [datetime]$Since,
    [string]$LogPath = ".\TeamsChatRedact_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# ============================================================================
#  EDIT THESE TWO TABLES TO TASTE
# ============================================================================

# REPLACE MAP - ordered, most-specific first. Each pattern replaces the WHOLE
# word it matches. Matching is case-insensitive. These run only on messages
# that do NOT hit the delete list below.
$ReplaceMap = [ordered]@{
    '\b\w*(?:shit|hsit)\w*\b'  = 'sassafrass'
}

# DELETE LIST - if a message matches ANY of these, the WHOLE message is
# removed (delete wins over replace).
$DeleteList = @(
    'mandy'
)

# ============================================================================

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

Write-Host 'Connecting to Microsoft Graph (device-code sign-in as yourself)...' -ForegroundColor Cyan
# Chat.ReadWrite covers both read + softDelete + sending a new chat message.
Connect-MgGraph -Scopes 'Chat.ReadWrite','User.Read' -NoWelcome

$me   = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me'
$myId = $me.id
Write-Host "Authenticated as: $($me.displayName) <$($me.userPrincipalName)>" -ForegroundColor Green
Write-Host "Mode: $(if ($DryRun) {'DRY RUN - nothing changes'} else {'LIVE - messages WILL be deleted/reposted'})" -ForegroundColor $(if ($DryRun) {'Yellow'} else {'Red'})

if (-not $DryRun) {
    $confirm = Read-Host "Type DELETE to confirm the live run"
    if ($confirm -cne 'DELETE') { Write-Host 'Aborted.'; return }
}

$results = [System.Collections.Generic.List[object]]::new()
$stats   = @{ Chats=0; Scanned=0; Deleted=0; Redacted=0; Skipped=0; Errors=0 }

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

            # 1) delete list wins
            $delHit = $DeleteList | Where-Object { $plain -match $_ } | Select-Object -First 1
            if ($delHit) {
                $entry = [pscustomobject]@{
                    Timestamp=$msg.createdDateTime; Chat=$label; Action=''; MatchedOn=$delHit
                    Before=$plain; After='(message removed)'; ChatId=$chat.id; MessageId=$msg.id; NewMessageId=''
                }
                if ($DryRun) { $entry.Action='WOULD DELETE'; Write-Host "  DELETE  <= [$delHit] $($plain.Substring(0,[Math]::Min(80,$plain.Length)))" -ForegroundColor Yellow }
                else {
                    try {
                        Invoke-GraphWithRetry POST "https://graph.microsoft.com/v1.0/me/chats/$($chat.id)/messages/$($msg.id)/softDelete" | Out-Null
                        $entry.Action='DELETED'; $stats.Deleted++; Write-Host "  DELETED <= [$delHit]" -ForegroundColor Red; Start-Sleep -Milliseconds 300
                    } catch { $entry.Action="ERROR: $($_.Exception.Message)"; $stats.Errors++ }
                }
                $results.Add($entry); continue
            }

            # 2) replace map -> if changed, delete original + repost cleaned
            $newHtml = $msg.body.content
            $matched = @()
            foreach ($k in $ReplaceMap.Keys) {
                if ($newHtml -match $k) { $matched += $k }
                $newHtml = [regex]::Replace($newHtml, $k, $ReplaceMap[$k], 'IgnoreCase')
            }
            if ($matched.Count -gt 0 -and $newHtml -ne $msg.body.content) {
                $newPlain = Get-PlainText $newHtml
                $entry = [pscustomobject]@{
                    Timestamp=$msg.createdDateTime; Chat=$label; Action=''; MatchedOn=($matched -join '; ')
                    Before=$plain; After=$newPlain; ChatId=$chat.id; MessageId=$msg.id; NewMessageId=''
                }
                if ($DryRun) { $entry.Action='WOULD REDACT'; Write-Host "  REDACT  '$($plain.Substring(0,[Math]::Min(60,$plain.Length)))' -> '$($newPlain.Substring(0,[Math]::Min(60,$newPlain.Length)))'" -ForegroundColor Yellow }
                else {
                    try {
                        # delete original
                        Invoke-GraphWithRetry POST "https://graph.microsoft.com/v1.0/me/chats/$($chat.id)/messages/$($msg.id)/softDelete" | Out-Null
                        # repost cleaned
                        $body = @{ body = @{ contentType='html'; content=$newHtml } }
                        $new  = Invoke-GraphWithRetry POST "https://graph.microsoft.com/v1.0/me/chats/$($chat.id)/messages" $body
                        $entry.Action='REDACTED'; $entry.NewMessageId=$new.id; $stats.Redacted++
                        Write-Host "  REDACTED (reposted as $($new.id))" -ForegroundColor Magenta; Start-Sleep -Milliseconds 400
                    } catch { $entry.Action="ERROR: $($_.Exception.Message)"; $stats.Errors++ }
                }
                $results.Add($entry)
            } else {
                $stats.Skipped++
            }
        }
        $msgUri = $page.'@odata.nextLink'
    }
}

if ($results.Count -gt 0) { $results | Export-Csv $LogPath -NoTypeInformation -Encoding UTF8; Write-Host "`nAudit log: $LogPath" -ForegroundColor Green }
Write-Host "`n===== SUMMARY =====" -ForegroundColor Cyan
Write-Host "Chats:    $($stats.Chats)"
Write-Host "Scanned:  $($stats.Scanned)  (my messages)"
Write-Host "Deleted:  $($stats.Deleted)"
Write-Host "Redacted: $($stats.Redacted)"
Write-Host "Errors:   $($stats.Errors)"
if ($DryRun) { Write-Host "`nDry run complete. Open the CSV, check the Before/After columns, then re-run without -DryRun." -ForegroundColor Yellow }
Disconnect-MgGraph | Out-Null
