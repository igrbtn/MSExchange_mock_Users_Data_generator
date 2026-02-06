# Exchange Mock Data Generator

PowerShell script that generates realistic email data on Microsoft Exchange Server 2019 for testing and migration validation. Creates bulk users with fully populated AD profiles, downloads/generates sample attachments, and sends thousands of emails (new messages, replies, forwards) with HTML formatting, inline images, and multi-language content.

## Architecture

```
┌──────────────────────────────────────────────────────────────────────┐
│                     Generate-ExchangeMockData.ps1                    │
├──────────────────────────────────────────────────────────────────────┤
│                                                                      │
│  Phase 1: Create Users + Populate AD Profiles                        │
│  ┌──────────────┐    ┌──────────────┐    ┌───────────────────────┐  │
│  │ New-Mailbox   │───▶│ OU=MockUsers │───▶│ users_credentials.csv │  │
│  │ x 300 users   │    │ (on PDC DC)  │    │ (full profile data)   │  │
│  └──────┬───────┘    └──────────────┘    └───────────────────────┘  │
│         │                                                            │
│         ▼                                                            │
│  ┌──────────────────────────────────────────────────────┐           │
│  │ Set-User: populate AD fields per user                 │           │
│  │  ┌────────────┐ ┌───────────┐ ┌────────────────────┐ │           │
│  │  │ Department │ │ Job Title │ │ Office / Building   │ │           │
│  │  │ Company    │ │ Manager   │ │ City / Country      │ │           │
│  │  │ Phone      │ │ Mobile    │ │ Street / PostalCode │ │           │
│  │  │ Initials   │ │ Notes     │ │ Description         │ │           │
│  │  └────────────┘ └───────────┘ └────────────────────┘ │           │
│  └──────────────────────────────────────────────────────┘           │
│                                                                      │
│  Phase 2: Configure SMTP + Throttling                                │
│  ┌───────────────────┐  ┌─────────────────────────────────────────┐ │
│  │ Detect Exchange    │  │ Auto-create ThrottlingPolicy            │ │
│  │ servers → user     │  │ "MockDataBulkSend" (Unlimited rates)   │ │
│  │ selects one        │  │ Apply to all 300 users                  │ │
│  └────────┬──────────┘  └─────────────────────────────────────────┘ │
│           │                                                          │
│           ▼                                                          │
│  ┌───────────────────┐  ┌─────────────────────────────────────────┐ │
│  │ Test auth SMTP     │  │ Raise transport delivery limits         │ │
│  │ port 465 + SSL     │  │ MaxConcurrentMailboxDeliveries = 100   │ │
│  └───────────────────┘  │ MaxConcurrentMailboxSubmissions = 100   │ │
│                          │ Receive Connector = unlimited           │ │
│                          └─────────────────────────────────────────┘ │
│                                                                      │
│  Phase 3: Prepare Attachments                                        │
│  ┌────────────┐  ┌────────────┐  ┌────────────┐                    │
│  │ 50x JPG    │  │ 30x TXT    │  │ 20x RTF    │                    │
│  │ (download  │  │ (multi-    │  │ (formatted │                    │
│  │  or gen)   │  │  language)  │  │  content)  │                    │
│  └────────────┘  └────────────┘  └────────────┘                    │
│        └──────────────┼──────────────┘                               │
│                       ▼                                              │
│              Attachments/ pool                                       │
│              (100 files total)                                       │
│                                                                      │
│  Phase 4: Send Emails (RunspacePool — parallel)                      │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │  Main thread (work generator)                                 │   │
│  │  ┌─────────┐                                                  │   │
│  │  │ Generate │── chunk of 100 work items ──┐                   │   │
│  │  │ emails   │                              │                   │   │
│  │  └────┬────┘                              ▼                   │   │
│  │       │                    ┌───────────────────────┐          │   │
│  │       │                    │   RunspacePool (N)    │          │   │
│  │       │                    │  ┌───┐┌───┐┌───┐     │          │   │
│  │       │                    │  │ W1││ W2││...│ x N  │          │   │
│  │       │                    │  └─┬─┘└─┬─┘└─┬─┘     │          │   │
│  │       │                    └────┼────┼────┼───────┘          │   │
│  │       │                         │    │    │                   │   │
│  │       │                         ▼    ▼    ▼                   │   │
│  │       │                    ┌───────────────────┐              │   │
│  │       │                    │  SMTP 465 + SSL   │              │   │
│  │       │                    │  (per-user auth)  │              │   │
│  │       │                    └─────────┬─────────┘              │   │
│  │       │                              │                        │   │
│  │       │◀── collect results ──────────┘                        │   │
│  │       │                                                       │   │
│  │  ┌────▼────┐                                                  │   │
│  │  │ Update   │ state.json, threads.json                        │   │
│  │  │ counters │ (resumable on re-run)                           │   │
│  │  └─────────┘                                                  │   │
│  └──────────────────────────────────────────────────────────────┘   │
│                                                                      │
│  Email Distribution:                                                 │
│   50% New messages ──▶ saved in threads.json (Message-ID)           │
│   30% Replies ───────▶ In-Reply-To + References headers             │
│   20% Forwards ──────▶ FW: prefix + original body quoted            │
│                                                                      │
│  Phase 5: Report                                                     │
│  ┌──────────────────────────────────────────────────────┐           │
│  │ DB size, mailbox stats, send rate, elapsed time      │           │
│  │ ──▶ generation_report.csv                            │           │
│  └──────────────────────────────────────────────────────┘           │
│                                                                      │
└──────────────────────────────────────────────────────────────────────┘
```

## User Profiles

Each mock user gets a fully populated Active Directory profile:

```
┌──────────────────────────────────────────────────────────────┐
│  Mock User: mockuser042                                      │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  Display Name:   Hiroshi Mueller                      │   │
│  │  First Name:     Hiroshi                              │   │
│  │  Last Name:      Mueller                              │   │
│  │  Initials:       HM                                   │   │
│  │  UPN:            mockuser042@lab.contoso.com          │   │
│  ├──────────────────────────────────────────────────────┤   │
│  │  Job Title:      Senior Analyst                       │   │
│  │  Department:     Research & Development               │   │
│  │  Company:        TechVista Corp                       │   │
│  │  Office:         Tower 1, Floor 10                    │   │
│  ├──────────────────────────────────────────────────────┤   │
│  │  City:           Tokyo                                │   │
│  │  Country:        JP                                   │   │
│  │  Street Address: 350 Enterprise Blvd                  │   │
│  │  Postal Code:    100-0001                             │   │
│  ├──────────────────────────────────────────────────────┤   │
│  │  Phone:          +1-555-4827                          │   │
│  │  Mobile:         +81-90-3847562                       │   │
│  │  Notes:          Mock user for testing purposes       │   │
│  └──────────────────────────────────────────────────────┘   │
└──────────────────────────────────────────────────────────────┘

Name pools: 300 first names + 300 last names
            from 12 cultures (EN, RU, ES/PT, FR, DE, CN, JP, AR, IN, KR, IT, Nordic)

Randomized fields (22 departments, 30 job titles, 25 offices,
                    26 cities/countries, 20 companies, 16 street addresses)
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
1. Create 300 users with mailboxes and full AD profiles in `OU=MockUsers`
2. Let you select which Exchange server to use for SMTP
3. Auto-configure throttling policies (user rate + transport delivery limits)
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
├── users_credentials.csv           # User accounts + passwords + profile data
├── state.json                      # Progress tracking (resumable)
├── threads.json                    # Message-ID tracking for threading
├── generation_report.csv           # Final statistics
└── generation_*.log                # Timestamped log files
```

## CSV Output Fields

The `users_credentials.csv` includes all profile data per user:

| Field | Example |
|-------|---------|
| Number | 42 |
| Alias | mockuser042 |
| UPN | mockuser042@lab.contoso.com |
| DisplayName | Hiroshi Mueller |
| Password | xK7#mPq2wR5! |
| SamAccountName | mockuser042 |
| FirstName | Hiroshi |
| LastName | Mueller |
| Department | Research & Development |
| Title | Senior Analyst |
| Office | Tower 1, Floor 10 |
| City | Tokyo |
| Country | JP |
| Company | TechVista Corp |
| Phone | +1-555-4827 |
| Mobile | +81-90-3847562 |

## How It Works

### Phase 1 — User Creation + AD Profile Population
Creates N mailbox-enabled users (`mockuser001` through `mockuser300`) in a dedicated OU. All AD and Exchange operations are pinned to the PDC Emulator to avoid replication lag. Names are drawn from 300 first names and 300 last names across 12 cultures (English, Russian, Spanish, Portuguese, French, German, Chinese, Japanese, Arabic, Indian, Korean, Italian, Nordic) and shuffled for unique combinations. After creating the mailbox, `Set-User` populates 12+ AD fields: Department, Title, Office, Company, City, Country, Street Address, Postal Code, Phone, Mobile, Initials, and Notes. Credentials and all profile data exported to CSV.

### Phase 2 — SMTP + Throttling Configuration
Enumerates Exchange servers and lets you pick which one to use for SMTP. Tests authenticated SMTP (SSL on port 465) with a real test message. Automatically:
- Creates a `MockDataBulkSend` throttling policy (MessageRateLimit=Unlimited, RecipientRateLimit=Unlimited)
- Applies it to all mock users
- Raises `MessageRateLimit` on Receive Connectors to unlimited
- Sets `MaxConcurrentMailboxDeliveries=100` and `MaxConcurrentMailboxSubmissions=100` on Transport and Mailbox Transport services (prevents `4.3.2` errors)

### Phase 3 — Attachments
Tries to download 50 JPEG images from the internet (picsum.photos, placehold.co, dummyimage.com). Falls back to generating gradient images with shapes using `System.Drawing` if the server has no internet access. Generates 30 multi-language text files and 20 RTF files with formatted content programmatically.

### Phase 4 — Email Sending (Parallel)
Uses a **PowerShell RunspacePool** for parallel SMTP delivery:
- Main thread generates work items (pre-builds HTML body, picks recipients/attachments)
- Dispatches chunks of 100 items to a pool of N worker threads
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

# Check user profile fields
Get-User mockuser001 | FL DisplayName,Title,Department,Office,City,Company,Phone

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

# Reset transport limits (optional)
Get-TransportService | Set-TransportService -MaxConcurrentMailboxDeliveries 20 -MaxConcurrentMailboxSubmissions 20
```

## Estimated Runtime

| Phase | Duration (300 users, 100GB) |
|-------|----------------------------|
| 1 — Create users + profiles | ~20 min |
| 2 — SMTP + throttling setup | ~5 min |
| 3 — Attachments | ~5-10 min |
| 4 — Send emails (10 threads) | ~4-8 hours |
| **Total** | **~5-9 hours** |

Performance scales roughly linearly with `-Threads`. With 20 threads: ~2-4 hours.

## License

MIT
