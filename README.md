# Redact-MyTeamsChatMessages.ps1

Bulk-clean **your own** Microsoft Teams chat messages via Microsoft Graph. The script scans every 1:1, group, and meeting chat you're a member of and applies two rules to messages **you** sent:

| Rule | Trigger | Result |
|------|---------|--------|
| **DELETE** | Message matches any pattern in `$DeleteList` | Whole message is soft-deleted |
| **REDACT** | Message matches any pattern in `$ReplaceMap` (and no delete hit) | Original is soft-deleted, a cleaned copy is reposted by you |

Delete always wins over replace.

---

## ⚠️ Read this before running

Microsoft Graph **cannot edit the body of a sent message**. "Redact" is therefore implemented as *soft-delete + repost*. Consequences:

- Recipients see the **"This message was deleted"** placeholder **and** the new cleaned message.
- The repost gets a **new timestamp** and lands at the **bottom of the chat** — not the original position.
- **Reactions and replies** attached to the original are lost.
- Soft delete is cosmetic only. **Compliance/eDiscovery copies, retention holds, and audit logs are unaffected.** This is not a tool for evading retention policy — don't use it as one.
- Only messages sent by the authenticated user are ever touched (delegated auth + sender-ID check).

---

## Requirements

- PowerShell 5.1+ or PowerShell 7
- Microsoft Graph PowerShell SDK (authentication module only):

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

- Delegated Graph permissions: `Chat.ReadWrite`, `User.Read` (requested at sign-in; consent may require admin approval depending on tenant policy)

## Authentication

The script uses interactive **device-code sign-in as yourself** (`Connect-MgGraph`). No app registration, client secret, or admin app is needed — but everything runs under your identity and shows up in the tenant audit log accordingly.

---

## Configuration

Edit the two tables near the top of the script:

### `$ReplaceMap` — redact (delete + repost cleaned)

Ordered hashtable of regex → replacement, most-specific first. Case-insensitive, applied to the raw HTML body.

```powershell
$ReplaceMap = [ordered]@{
    '\b\w*(?:shit|hsit)\w*\b'  = 'sassafrass'
}
```

### `$DeleteList` — remove the whole message

Array of regex patterns. Any match deletes the entire message.

```powershell
$DeleteList = @(
    'mandy'
)
```

> Patterns are regex, matched against the **plain-text** version of the message for deletes and the **HTML** for replaces. Test patterns carefully — `'mandy'` also matches "Normandy".

---

## Usage

```powershell
# 1. ALWAYS dry-run first — nothing changes, output goes to console + CSV
.\Redact-MyTeamsChatMessages.ps1 -DryRun

# 2. Review the CSV (Before/After columns), then run live
.\Redact-MyTeamsChatMessages.ps1
```

A live run prompts you to type `DELETE` before touching anything.

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-DryRun` | switch | off | Report what *would* happen; no writes |
| `-Since` | datetime | (none) | Only process messages created on/after this date. Older messages stop pagination per chat |
| `-LogPath` | string | `.\TeamsChatRedact_<timestamp>.csv` | Where the audit CSV is written |

### Examples

```powershell
# Only touch messages from the last 30 days
.\Redact-MyTeamsChatMessages.ps1 -DryRun -Since (Get-Date).AddDays(-30)

# Custom log location
.\Redact-MyTeamsChatMessages.ps1 -LogPath 'C:\Logs\teams_redact.csv'
```

---

## Output

**Console:** per-chat progress, per-message DELETE/REDACT actions, and a summary block (chats, scanned, deleted, redacted, errors).

**CSV audit log** with one row per actioned message:

| Column | Meaning |
|--------|---------|
| `Timestamp` | Original message created time |
| `Chat` | Chat topic, or the other members' names for untitled chats |
| `Action` | `DELETED`, `REDACTED`, `WOULD DELETE`, `WOULD REDACT`, or `ERROR: …` |
| `MatchedOn` | The pattern(s) that triggered the action |
| `Before` / `After` | Plain-text message before and after |
| `ChatId` / `MessageId` / `NewMessageId` | Graph IDs (NewMessageId populated on repost) |

## Behavior notes

- **Throttling:** Graph 429 responses honor `Retry-After` (up to 6 retries); 5xx errors back off up to 4 retries.
- **Pacing:** short sleeps (300–400 ms) after each write to stay under chat-message rate limits.
- **Skipped:** system messages, already-deleted messages, and messages from anyone other than you.
- **Scope:** iterates *all* chats you're in via `/me/chats`; use `-Since` to bound large histories.

## Known limitations

- Cannot preserve original message position, timestamp, reactions, or thread replies on redacted messages.
- Channel (team) messages are **not** processed — chats only (`/me/chats`).
- Delete patterns run against plain text while replace patterns run against HTML, so a replace pattern spanning an HTML tag boundary may not match.
- Reposting fires notifications — a large live run will ping people.
