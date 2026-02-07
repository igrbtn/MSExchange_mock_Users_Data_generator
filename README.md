# Exchange Mock Data Generator

PowerShell script that generates realistic email data on Microsoft Exchange Server 2019 for testing and migration validation. Creates bulk users with fully populated AD profiles, populates contacts and calendar events, downloads/generates sample attachments, and sends thousands of emails (new messages, replies, forwards) with HTML formatting, inline images, CC recipients, and multi-language content.

## Architecture

```
┌──────────────────────────────────────────────────────────────────────┐
│                     Generate-ExchangeMockData.ps1                    │
├──────────────────────────────────────────────────────────────────────┤
│                                                                      │
│  Phase 1: Create Users + Populate AD Profiles                        │
│  ┌────────────────────────────────────────────────────────────────┐ │
│  │ Interactive setup prompts:                                      │ │
│  │  "How many users to create? [300]"                              │ │
│  │  "OU name for mock users? [MockUsers]"                          │ │
│  │  "Password mode: [1] Random per user  [2] Same for all"        │ │
│  │  "Available Mailbox Databases:"                                 │ │
│  │   [1] MDB01 (EX01, 52 GB)                                      │ │
│  │   [2] MDB02 (EX02, 18 GB)                                      │ │
│  └────────────────────────────────────────────────────────────────┘ │
│  ┌──────────────┐    ┌──────────────┐    ┌───────────────────────┐  │
│  │ New-Mailbox   │───▶│ OU=MockUsers │───▶│ users_credentials.csv │  │
│  │ x N users     │    │ (on PDC DC)  │    │ (full profile data)   │  │
│  └──────┬───────┘    └──────────────┘    └───────────────────────┘  │
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
│  Phase 2: Configure SMTP + EWS + Throttling                         │
│  ┌────────────────────────────────────────────────────────────────┐ │
│  │ Interactive server selection (shown once, chosen separately):   │ │
│  │                                                                 │ │
│  │  Available Exchange servers:                                    │ │
│  │   [1] EX01.lab.contoso.com  (Mailbox)                          │ │
│  │   [2] EX02.lab.contoso.com  (Mailbox)                          │ │
│  │                                                                 │ │
│  │  Select SMTP server number (1-2): _                            │ │
│  │  Select EWS server number (1-2): _                             │ │
│  └────────────────────────────────────────────────────────────────┘ │
│  ┌───────────────────┐  ┌─────────────────────────────────────────┐ │
│  │ Test auth SMTP     │  │ Auto-create ThrottlingPolicy            │ │
│  │ port 465 + SSL     │  │ "MockDataBulkSend" (Unlimited rates)   │ │
│  └───────────────────┘  │ Apply to all 300 users                  │ │
│                          ├─────────────────────────────────────────┤ │
│                          │ Raise transport delivery limits         │ │
│                          │ MaxConcurrentMailboxDeliveries = 100    │ │
│                          │ MaxConcurrentMailboxSubmissions = 100   │ │
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
│  Phase 4: Provision Folders → Contacts → Calendar → Send Emails      │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │                                                               │   │
│  │  4.0  Provision Default Folders (EWS)                         │   │
│  │  ┌────────────────────────────────────────────────────────┐  │   │
│  │  │ Auth EWS GetFolder(sentitems) per user (parallel)      │  │   │
│  │  │ → Exchange creates Inbox, Sent Items, Drafts, etc.     │  │   │
│  │  └────────────────────────────────────────────────────────┘  │   │
│  │                                                               │   │
│  │  4.1  Create Contacts (EWS)                                   │   │
│  │  ┌────────────────────────────────────────────────────────┐  │   │
│  │  │ Each user gets 10-30 contacts from other mock users    │  │   │
│  │  │ Fields: Name, Email, Company, Phone, Title, Department │  │   │
│  │  │ Created via EWS CreateItem in Contacts folder          │  │   │
│  │  └────────────────────────────────────────────────────────┘  │   │
│  │                                                               │   │
│  │  4.2  Calendar Events (SMTP + iCalendar)                      │   │
│  │  ┌────────────────────────────────────────────────────────┐  │   │
│  │  │ Each user creates 5-15 meetings (METHOD:REQUEST)       │  │   │
│  │  │ 2-8 attendees, spread across -30 to +90 days           │  │   │
│  │  │ 30 meeting subjects, 17 locations, 15-min reminders    │  │   │
│  │  │ Sent as text/calendar MIME → Exchange creates events   │  │   │
│  │  └────────────────────────────────────────────────────────┘  │   │
│  │                                                               │   │
│  │  4a-c Send Emails (RunspacePool — parallel)                   │   │
│  │  ┌────────────────────────────────────────────────────────┐  │   │
│  │  │  Main thread generates chunks → RunspacePool (N)       │  │   │
│  │  │  ┌───────────────────────────────────────────────────┐ │  │   │
│  │  │  │  4a: 50% New messages  (1-5 To + 0-4 CC)         │ │  │   │
│  │  │  │  4b: 30% Replies       (To + 0-3 CC)             │ │  │   │
│  │  │  │  4c: 20% Forwards      (1-3 To + 0-3 CC)         │ │  │   │
│  │  │  └───────────────────────────────────────────────────┘ │  │   │
│  │  │  SMTP 465 + SSL, per-user auth, ~110 email subjects    │  │   │
│  │  │  Results → state.json + threads.json (resumable)       │  │   │
│  │  └────────────────────────────────────────────────────────┘  │   │
│  └──────────────────────────────────────────────────────────────┘   │
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
│  │  Display Name:   MockOrg Mueller Hiroshi               │   │
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

## Contacts & Calendar

```
┌──────────────────────────────────────────────────────────────┐
│  Contacts (per user)                                         │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  10-30 contacts created in each user's mailbox        │   │
│  │  Sourced from other mock users' AD profiles           │   │
│  │                                                        │   │
│  │  Fields populated:                                     │   │
│  │   • First Name / Last Name / Display Name              │   │
│  │   • Email Address (EmailAddress1)                      │   │
│  │   • Company Name                                       │   │
│  │   • Business Phone                                     │   │
│  │   • Job Title                                          │   │
│  │   • Department                                         │   │
│  │                                                        │   │
│  │  Created via EWS CreateItem → Contacts folder          │   │
│  └──────────────────────────────────────────────────────┘   │
│                                                              │
│  Calendar Events (per user)                                  │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  5-15 meeting invitations per organizer               │   │
│  │  2-8 attendees per meeting                            │   │
│  │  Time range: -30 days (past) to +90 days (future)     │   │
│  │  Duration: 30 / 45 / 60 / 90 / 120 min               │   │
│  │  Business hours: 08:00 — 18:00                        │   │
│  │                                                        │   │
│  │  30 meeting subjects (EN, RU, FR, PT)                  │   │
│  │  17 locations (rooms, online, buildings)               │   │
│  │  15-minute reminder alarm                              │   │
│  │                                                        │   │
│  │  Sent as iCalendar METHOD:REQUEST via SMTP             │   │
│  │  → Exchange creates calendar items for all attendees   │   │
│  └──────────────────────────────────────────────────────┘   │
└──────────────────────────────────────────────────────────────┘
```

## Email Content

```
┌───────────────────────────────────────────────────────────┐
│  5 HTML Templates         10 Languages    ~110 Subjects   │
│  ┌──────────────────┐  ┌────────────────┐  ┌───────────┐ │
│  │ Business Formal   │  │ English        │  │ EN (64)   │ │
│  │ Casual             │  │ Russian        │  │ RU (10)   │ │
│  │ Report (tables)    │  │ Spanish        │  │ FR (8)    │ │
│  │ Newsletter (img)   │  │ French         │  │ ES (7)    │ │
│  │ Simple Reply       │  │ German         │  │ DE (8)    │ │
│  └──────────────────┘  │ Chinese        │  │ JP (6)    │ │
│                         │ Japanese       │  │ CN (6)    │ │
│  Recipients per email:  │ Arabic         │  │ PT (6)    │ │
│   To:  1-5 recipients   │ Portuguese     │  │ IT (6)    │ │
│   CC:  40% chance 1-4   │ Italian        │  │ KR (5)    │ │
│                         └────────────────┘  │ AR (5)    │ │
│  Attachments per email:                      └───────────┘ │
│   40% — none                                               │
│   30% — 1 small file                                       │
│   20% — 1 medium file                                      │
│   10% — 1-3 large files                                    │
│                                                             │
│  30% of emails include inline images (CID)                 │
│  All emails have HTML signatures                           │
│  Replies: In-Reply-To + References headers + quoted text   │
│  Forwards: FW: prefix + original sender/subject in body    │
└───────────────────────────────────────────────────────────┘
```

## Requirements

- **Exchange Server 2019** (on-premises)
- **Windows PowerShell 5.1** (runs in Exchange Management Shell)
- **Permissions**: Domain Admin + Organization Management role group
- **SMTP**: Port 465 (SSL) must be accessible on the target Exchange server
- **EWS**: HTTPS (443) for folder provisioning, contacts, and calendar
- **ActiveDirectory** PowerShell module (RSAT)

## Quick Start

```powershell
# Copy to Exchange server, then run from Exchange Management Shell:
.\Generate-ExchangeMockData.ps1
```

The script will interactively ask you to configure:
1. **User count** — how many mock users to create (default: 300)
2. **OU name** — organizational unit for mock accounts (default: MockUsers)
3. **Password mode** — random per user or same password for all
4. **Mailbox Database** — which database to create user mailboxes in
5. **SMTP Server** — which Exchange server to use for sending emails
6. **EWS Server** — which Exchange server to use for contacts/calendar/folders

Then it:
7. Creates users with mailboxes and full AD profiles
5. Auto-configures throttling policies (user rate + transport delivery limits)
6. Downloads/generates 100 sample attachment files
7. Provisions default mailbox folders (Sent Items, Drafts, etc.)
8. Creates 10-30 contacts per user from the mock user pool
9. Sends 5-15 calendar meeting invitations per user
10. Sends ~68,000 emails in parallel (10 threads) to reach ~100GB
11. Prints a summary report

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-TargetSizeGB` | `100` | Target total mailbox database size |
| `-UserCount` | interactive | Number of mock users (prompts if not set, default 300) |
| `-MockUsersOU` | interactive | OU name for mock accounts (prompts if not set, default MockUsers) |
| `-UserPassword` | interactive | Password for all users; if empty, prompts for mode (random/same) |
| `-Database` | interactive | Mailbox database name for user creation |
| `-SmtpServer` | interactive | Exchange server FQDN for SMTP sending |
| `-EwsServer` | interactive | Exchange server FQDN for EWS operations |
| `-SmtpPort` | `465` | SMTP port (465 for SSL, or 587) |
| `-Threads` | `10` | Parallel send/provision threads |
| `-ChunkSize` | `100` | Work items per batch |
| `-SkipToPhase` | auto | Resume from a specific phase (1-5) |
| `-UserPrefix` | `mockuser` | Username prefix (e.g. mockuser001) |
| `-PasswordLength` | `12` | Length of random passwords |
| `-InlineImagePercent` | `30` | % of emails with inline images |
| `-Force` | off | Skip confirmation prompts |

## Examples

```powershell
# Default: 300 users, 100GB, 10 threads (interactive server/DB selection)
.\Generate-ExchangeMockData.ps1

# Smaller test: 50 users, 10GB (skip user count prompt)
.\Generate-ExchangeMockData.ps1 -UserCount 50 -TargetSizeGB 10

# Same password for all users (skip password prompt)
.\Generate-ExchangeMockData.ps1 -UserPassword "P@ssw0rd123!"

# Pre-select everything (no interactive prompts at all)
.\Generate-ExchangeMockData.ps1 -UserCount 300 -MockUsersOU "MockUsers" -UserPassword "P@ssw0rd123!" -Database "MDB01" -SmtpServer ex01.contoso.com -EwsServer ex01.contoso.com

# Resume from Phase 4 with more threads
.\Generate-ExchangeMockData.ps1 -SkipToPhase 4 -Threads 20

# Use different servers for SMTP and EWS
.\Generate-ExchangeMockData.ps1 -SmtpServer ex01.contoso.com -EwsServer ex02.contoso.com

# Re-run Phase 2 (re-select servers, re-apply throttling)
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
| DisplayName | MockOrg Mueller Hiroshi |
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
Prompts for user count (default 300), OU name (default MockUsers), and password mode (random per user or same for all). Then interactively asks which mailbox database to use (shows all databases with server and size info). Creates N mailbox-enabled users (`mockuser001` through `mockuserN`) in the dedicated OU. All AD and Exchange operations are pinned to the PDC Emulator to avoid replication lag. Names are drawn from 300 first names and 300 last names across 12 cultures (English, Russian, Spanish, Portuguese, French, German, Chinese, Japanese, Arabic, Indian, Korean, Italian, Nordic) and shuffled for unique combinations. After creating the mailbox, `Set-User` populates 12+ AD fields: Department, Title, Office, Company, City, Country, Street Address, Postal Code, Phone, Mobile, Initials, and Notes. Credentials and all profile data exported to CSV.

### Phase 2 — SMTP + EWS + Throttling Configuration
Enumerates Exchange servers and lets you choose separately which server to use for **SMTP** (email sending) and **EWS** (folder provisioning, contacts, calendar). Tests authenticated SMTP (SSL on port 465) with a real test message. Displays an infrastructure summary. Automatically:
- Creates a `MockDataBulkSend` throttling policy (MessageRateLimit=Unlimited, RecipientRateLimit=Unlimited)
- Applies it to all mock users
- Raises `MessageRateLimit` on Receive Connectors to unlimited
- Sets `MaxConcurrentMailboxDeliveries=100` and `MaxConcurrentMailboxSubmissions=100` on Transport and Mailbox Transport services (prevents `4.3.2` errors)

### Phase 3 — Attachments
Tries to download 50 JPEG images from the internet (picsum.photos, placehold.co, dummyimage.com). Falls back to generating gradient images with shapes using `System.Drawing` if the server has no internet access. Generates 30 multi-language text files and 20 RTF files with formatted content programmatically.

### Phase 4 — Folders, Contacts, Calendar, Emails

**4.0 Folder Provisioning** — Authenticates as each user via EWS `GetFolder(sentitems)` to force Exchange to create all default folders (Inbox, Sent Items, Drafts, Deleted Items, etc.). Runs in parallel via RunspacePool.

**4.1 Contacts** — Creates 10-30 contacts in each user's Contacts folder via EWS `CreateItem`. Contact data is sourced from other mock users' AD profiles (name, email, company, phone, job title, department). Runs in parallel.

**4.2 Calendar Events** — Each user organizes 5-15 meetings by sending iCalendar `METHOD:REQUEST` invitations via SMTP. Meetings have 2-8 attendees, span -30 to +90 days, use business hours, and include 15-minute reminders. Exchange processes the iCalendar and creates proper calendar items for all attendees. Runs in parallel.

**4a-c Email Sending** — Uses a PowerShell RunspacePool for parallel SMTP delivery:
- Main thread generates work items (pre-builds HTML body, picks recipients/attachments/CC)
- Dispatches chunks of 100 items to a pool of N worker threads
- Each worker authenticates as the sender user via SMTP 465 + SSL
- ~110 email subjects across 10 languages
- **To**: 1-5 recipients, **CC**: 40% chance of 1-4 additional recipients
- Replies include CC (30% chance), forwards go to 1-3 To + CC (35% chance)
- Results collected, counters updated, state saved after each chunk
- Fully resumable — tracks progress in `state.json`

Email distribution: 50% new messages, 30% replies (with `In-Reply-To`/`References` headers and quoted text), 20% forwards (with original sender info in body).

### Phase 5 — Report
Queries actual mailbox database size and sample mailbox statistics. Exports summary to CSV.

## State Tracking

The script saves progress to `state.json` after every step. Tracked fields include:

| Field | Purpose |
|-------|---------|
| `database` | Selected mailbox database |
| `smtpServer` / `smtpPort` | Selected SMTP server |
| `ewsServer` | Selected EWS server |
| `foldersProvisioned` | Default folders created |
| `contactsCreated` | Contacts populated |
| `calendarCreated` | Calendar events sent |
| `emailsSent` | Total emails sent |
| `estimatedSizeGB` | Estimated cumulative size |

If interrupted (Ctrl+C, reboot, error), simply re-run the script — it picks up from where it left off. Server selections are remembered from state.

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

# Check contacts (via EWS or OWA)
# Open OWA as a mock user (use credentials from CSV)
# https://mail.yourdomain.com/owa
```

## Cleanup

To remove all mock data:

```powershell
# Remove all mock mailboxes
Get-Mailbox -OrganizationalUnit MockUsers | Remove-Mailbox -Confirm:$false

# Remove the OU (replace DC=yourdomain,DC=com with your domain)
Remove-ADOrganizationalUnit "OU=MockUsers,DC=yourdomain,DC=com" -Recursive -Confirm:$false

# Remove throttling policy
Remove-ThrottlingPolicy -Identity MockDataBulkSend -Confirm:$false

# Reset transport limits (optional)
Get-TransportService | Set-TransportService -MaxConcurrentMailboxDeliveries 20 -MaxConcurrentMailboxSubmissions 20

# Clean up local state files
Remove-Item state.json, threads.json, users_credentials.csv -ErrorAction SilentlyContinue
```

## Download / Update on Server

```powershell
# First time — clone (if git is installed):
git clone https://github.com/igrbtn/MSExchange_mock_Users_Data_generator.git
cd MSExchange_mock_Users_Data_generator

# First time — without git (download zip):
Invoke-WebRequest -Uri "https://github.com/igrbtn/MSExchange_mock_Users_Data_generator/archive/refs/heads/main.zip" -OutFile "$env:TEMP\mockdata.zip"
Expand-Archive "$env:TEMP\mockdata.zip" -DestinationPath C:\Scripts -Force
Rename-Item "C:\Scripts\MSExchange_mock_Users_Data_generator-main" "C:\Scripts\ExchangeMockData"

# Pull updates (if cloned with git):
git pull
```

## Estimated Runtime

| Phase | Duration (300 users, 100GB) |
|-------|----------------------------|
| 1 — Create users + profiles | ~20 min |
| 2 — SMTP/EWS + throttling setup | ~5 min |
| 3 — Attachments | ~5-10 min |
| 4.0 — Folder provisioning | ~5 min |
| 4.1 — Contacts (~6,000 total) | ~10-15 min |
| 4.2 — Calendar (~3,000 meetings) | ~10-15 min |
| 4a-c — Send emails (10 threads) | ~4-8 hours |
| **Total** | **~5-9 hours** |

Performance scales roughly linearly with `-Threads`. With 20 threads: ~2-4 hours for email sending.

## License

MIT
