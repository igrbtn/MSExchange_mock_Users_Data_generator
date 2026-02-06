# Exchange Mock Data Generator

PowerShell script that generates realistic email data on Microsoft Exchange Server 2019 for testing and migration validation. Creates bulk users with mailboxes, downloads/generates sample attachments, and sends thousands of emails (new messages, replies, forwards) with HTML formatting, inline images, and multi-language content.

## Architecture

```
┌──────────────────────────────────────────────────────────────────┐
│                    Generate-ExchangeMockData.ps1                 │
├──────────────────────────────────────────────────────────────────┤
│                                                                  │
│  Phase 1: Create Users                                           │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────────┐   │
│  │ New-Mailbox   │───▶│ OU=MockUsers │───▶│ credentials.csv  │   │
│  │ x 300 users   │    │ (on PDC DC)  │    │ (UPN + password) │   │
│  └──────────────┘    └──────────────┘    └──────────────────┘   │
│                                                                  │
│  Phase 2: Configure SMTP + Throttling                            │
│  ┌──────────────────┐  ┌────────────────────────────────────┐   │
│  │ Detect Exchange   │  │ Auto-create ThrottlingPolicy       │   │
│  │ servers (select)  │  │ "MockDataBulkSend" (Unlimited)     │   │
│  └────────┬─────────┘  │ Apply to all 300 users             │   │
│           │             └────────────────────────────────────┘   │
│           ▼                                                      │
│  ┌──────────────────┐                                            │
│  │ Test auth SMTP    │                                           │
│  │ port 465 + SSL    │                                           │
│  └──────────────────┘                                            │
│                                                                  │
│  Phase 3: Prepare Attachments                                    │
│  ┌────────────┐  ┌────────────┐  ┌────────────┐                │
│  │ 50x JPG    │  │ 30x TXT    │  │ 20x RTF    │                │
│  │ (download  │  │ (multi-    │  │ (formatted │                │
│  │  or gen)   │  │  language)  │  │  content)  │                │
│  └────────────┘  └────────────┘  └────────────┘                │
│        └──────────────┼──────────────┘                           │
│                       ▼                                          │
│              Attachments/ pool                                   │
│              (100 files total)                                   │
│                                                                  │
│  Phase 4: Send Emails (RunspacePool — parallel)                  │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │  Main thread (work generator)                             │   │
│  │  ┌─────────┐                                              │   │
│  │  │ Generate │── chunk of 100 work items ──┐               │   │
│  │  │ emails   │                              │               │   │
│  │  └────┬────┘                              ▼               │   │
│  │       │                    ┌───────────────────────┐      │   │
│  │       │                    │   RunspacePool (10)   │      │   │
│  │       │                    │  ┌───┐┌───┐┌───┐     │      │   │
│  │       │                    │  │ W1││ W2││...│ x10  │      │   │
│  │       │                    │  └─┬─┘└─┬─┘└─┬─┘     │      │   │
│  │       │                    └────┼────┼────┼───────┘      │   │
│  │       │                         │    │    │               │   │
│  │       │                         ▼    ▼    ▼               │   │
│  │       │                    ┌───────────────────┐          │   │
│  │       │                    │  SMTP 465 + SSL   │          │   │
│  │       │                    │  (per-user auth)  │          │   │
│  │       │                    └─────────┬─────────┘          │   │
│  │       │                              │                    │   │
│  │       │◀── collect results ──────────┘                    │   │
│  │       │                                                   │   │
│  │  ┌────▼────┐                                              │   │
│  │  │ Update   │ state.json, threads.json                    │   │
│  │  │ counters │ (resumable)                                 │   │
│  │  └─────────┘                                              │   │
│  └──────────────────────────────────────────────────────────┘   │
│                                                                  │
│  Email Distribution:                                             │
│   50% New messages ──▶ saved in threads.json (Message-ID)       │
│   30% Replies ───────▶ In-Reply-To + References headers         │
│   20% Forwards ──────▶ FW: prefix + original body quoted        │
│                                                                  │
│  Phase 5: Report                                                 │
│  ┌──────────────────────────────────────────────────────┐       │
│  │ DB size, mailbox stats, send rate, elapsed time      │       │
│  │ ──▶ generation_report.csv                            │       │
│  └──────────────────────────────────────────────────────┘       │
│                                                                  │
└──────────────────────────────────────────────────────────────────┘
```

## Email Content

```
┌────────────────────────────────────────────┐
│  5 HTML Templates         10 Languages     │
│  ┌──────────────────┐  ┌────────────────┐  │
│  │ Business Formal   │  │ English        │  │
│  │ Casual             │  │ Russian        │  │
│  │ Report (tables)    │  │ Spanish        │  │
│  │ Newsletter (img)   │  │ French         │  │
│  │ Simple Reply       │  │ German         │  │
│  └──────────────────┘  │ Chinese        │  │
│                         │ Japanese       │  │
│  Attachments per email: │ Arabic         │  │
│   40% — none            │ Portuguese     │  │
│   30% — 1 small file    │ Italian        │  │
│   20% — 1 medium file   └────────────────┘  │
│   10% — 1-3 large files                     │
│                                              │
│  30% of emails include inline images (CID)  │
│  All emails have HTML signatures            │
└────────────────────────────────────────────┘
```

## Requirements

- **Exchange Server 2019** (on-premises)
- **Windows PowerShell 5.1** (runs in Exchange Management Shell)
- **Permissions**: Domain Admin + Organization Management role group
- **SMTP**: Port 465 (SSL) must be accessible on the target Exchange server
- **ActiveDirectory** PowerShell module (RSAT)

## Quick Start

```powershell
# Copy to Exchange server, then run from Exchange Management Shell:
.\Generate-ExchangeMockData.ps1
```

The script will:
1. Create 300 users with mailboxes in `OU=MockUsers`
2. Let you select which Exchange server to use for SMTP
3. Auto-configure throttling policy for unlimited send rate
4. Download/generate 100 sample attachment files
5. Send ~68,000 emails in parallel (10 threads) to reach ~100GB
6. Print a summary report

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-TargetSizeGB` | `100` | Target total mailbox database size |
| `-UserCount` | `300` | Number of mock users to create |
| `-Threads` | `10` | Parallel SMTP send threads |
| `-ChunkSize` | `100` | Work items per batch |
| `-SmtpServer` | auto-detect | Exchange server FQDN for SMTP |
| `-SmtpPort` | `465` | SMTP port (465 for SSL, or 587) |
| `-SkipToPhase` | auto | Resume from a specific phase (1-5) |
| `-UserPrefix` | `mockuser` | Username prefix (e.g. mockuser001) |
| `-MockUsersOU` | `MockUsers` | OU name for mock user accounts |
| `-InlineImagePercent` | `30` | % of emails with inline images |
| `-Force` | off | Skip confirmation prompts |

## Examples

```powershell
# Default: 300 users, 100GB, 10 threads
.\Generate-ExchangeMockData.ps1

# Smaller test: 50 users, 10GB
.\Generate-ExchangeMockData.ps1 -UserCount 50 -TargetSizeGB 10

# Resume from Phase 4 with more threads
.\Generate-ExchangeMockData.ps1 -SkipToPhase 4 -Threads 20

# Use a specific Exchange server and port 587
.\Generate-ExchangeMockData.ps1 -SmtpServer ex01.contoso.com -SmtpPort 587

# Re-run Phase 2 (re-select server, re-apply throttling)
Remove-Item state.json
.\Generate-ExchangeMockData.ps1 -SkipToPhase 2
```

## Files Generated

```
ExchangeMockData/
├── Generate-ExchangeMockData.ps1   # Main script
├── Attachments/                    # Sample files (auto-created)
│   ├── jpg/                        # 50 images (downloaded or generated)
│   ├── txt/                        # 30 multi-language text files
│   └── rtf/                        # 20 formatted RTF documents
├── users_credentials.csv           # User accounts + passwords
├── state.json                      # Progress tracking (resumable)
├── threads.json                    # Message-ID tracking for threading
├── generation_report.csv           # Final statistics
└── generation_*.log                # Timestamped log files
```

## How It Works

### Phase 1 — User Creation
Creates N mailbox-enabled users (`mockuser001` through `mockuser300`) in a dedicated OU. All AD and Exchange operations are pinned to the PDC Emulator to avoid replication lag issues. International display names (150 first + 150 last names across 15 cultures). Credentials exported to CSV.

### Phase 2 — SMTP Configuration
Enumerates Exchange servers and lets you pick which one to use. Tests authenticated SMTP (SSL on port 465) with a real test message. Automatically creates a `MockDataBulkSend` throttling policy with unlimited message/recipient rates and applies it to all mock users. Also raises `MessageRateLimit` on relevant Receive Connectors.

### Phase 3 — Attachments
Tries to download 50 JPEG images from the internet (picsum.photos, placehold.co, dummyimage.com). Falls back to generating gradient images with shapes using `System.Drawing` if the server has no internet access. Generates 30 multi-language text files and 20 RTF files with formatted content programmatically.

### Phase 4 — Email Sending (Parallel)
Uses a **PowerShell RunspacePool** for parallel SMTP delivery:
- Main thread generates work items (pre-builds HTML body, picks recipients/attachments)
- Dispatches chunks of 100 items to a pool of 10 worker threads
- Each worker authenticates as the sender user via SMTP 465 + SSL
- Results collected, counters updated, state saved after each chunk
- Fully resumable — tracks progress in `state.json`

Email distribution: 50% new messages, 30% replies (with `In-Reply-To`/`References` headers and quoted text), 20% forwards (with original sender info in body).

### Phase 5 — Report
Queries actual mailbox database size and sample mailbox statistics. Exports summary to CSV.

## Resumability

The script saves progress to `state.json` after every chunk. If interrupted (Ctrl+C, reboot, error), simply re-run the script — it picks up from where it left off.

To force a full restart:
```powershell
Remove-Item state.json, threads.json -ErrorAction SilentlyContinue
.\Generate-ExchangeMockData.ps1
```

## Verification

After the script completes:

```powershell
# Check total users
Get-Mailbox -OrganizationalUnit MockUsers | Measure-Object

# Check database size
Get-MailboxDatabase -Status | Select Name, DatabaseSize

# Check a sample mailbox
Get-MailboxStatistics mockuser001 | Select ItemCount, TotalItemSize

# Open OWA as a mock user (use credentials from CSV)
# https://mail.yourdomain.com/owa
```

## Cleanup

To remove all mock data:

```powershell
# Remove all mock mailboxes
Get-Mailbox -OrganizationalUnit MockUsers | Remove-Mailbox -Confirm:$false

# Remove the OU
Remove-ADOrganizationalUnit "OU=MockUsers,DC=yourdomain,DC=com" -Recursive -Confirm:$false

# Remove throttling policy
Remove-ThrottlingPolicy -Identity MockDataBulkSend -Confirm:$false
```

## Estimated Runtime

| Phase | Duration (300 users, 100GB) |
|-------|----------------------------|
| 1 — Create users | ~15 min |
| 2 — SMTP setup | ~5 min |
| 3 — Attachments | ~5-10 min |
| 4 — Send emails (10 threads) | ~4-8 hours |
| **Total** | **~5-9 hours** |

Performance scales roughly linearly with `-Threads`. With 20 threads: ~2-4 hours.

## License

MIT
