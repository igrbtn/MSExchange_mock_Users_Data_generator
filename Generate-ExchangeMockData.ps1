#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange Mock Data Generator — creates 300 users and ~100GB of realistic email data.

.DESCRIPTION
    Runs on Exchange Server 2019 with Domain Admin + Organization Management.
    Phase 1: Create 300 mailbox-enabled users
    Phase 2: Validate SMTP relay on localhost:25
    Phase 3: Download/generate sample attachments (JPG, TXT, RTF)
    Phase 4: Send ~34,000 emails via authenticated SMTP (new, reply, forward) with HTML formatting and inline images
    Phase 5: Generate report

    Resumable via state.json — re-run the script to continue from where it left off.

.PARAMETER SkipToPhase
    Resume from a specific phase (1-5). By default, auto-detects from state.json.

.PARAMETER TargetSizeGB
    Target total mailbox database size in GB. Default: 100.

.PARAMETER UserCount
    Number of mock users to create. Default: 300.

.EXAMPLE
    .\Generate-ExchangeMockData.ps1
    .\Generate-ExchangeMockData.ps1 -TargetSizeGB 50 -UserCount 100
    .\Generate-ExchangeMockData.ps1 -SkipToPhase 4
#>

param(
    [int]$SkipToPhase = 0,
    [int]$TargetSizeGB = 100,
    [int]$UserCount = 300,
    [string]$UserPrefix = "mockuser",
    [string]$MockUsersOU = "MockUsers",
    [int]$PasswordLength = 12,
    [int]$Threads = 10,
    [int]$ChunkSize = 100,
    [int]$MaxAttachmentSizeMB = 10,
    [int]$InlineImagePercent = 30,
    [string]$SmtpServer = "",
    [int]$SmtpPort = 465,
    [switch]$Force
)

$ErrorActionPreference = "Continue"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$AttachmentsDir = Join-Path $ScriptDir "Attachments"
$JpgDir = Join-Path $AttachmentsDir "jpg"
$TxtDir = Join-Path $AttachmentsDir "txt"
$RtfDir = Join-Path $AttachmentsDir "rtf"
$StateFile = Join-Path $ScriptDir "state.json"
$ThreadsFile = Join-Path $ScriptDir "threads.json"
$CredsFile = Join-Path $ScriptDir "users_credentials.csv"
$ReportFile = Join-Path $ScriptDir "generation_report.csv"
$LogFile = Join-Path $ScriptDir "generation_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

#####################################################################
# HELPER FUNCTIONS
#####################################################################

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Message"
    Write-Host $line -ForegroundColor $(switch ($Level) {
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        "OK"    { "Green" }
        default { "White" }
    })
    Add-Content -Path $LogFile -Value $line -ErrorAction SilentlyContinue
}

function Get-State {
    if (Test-Path $StateFile) {
        return Get-Content $StateFile -Raw | ConvertFrom-Json
    }
    return @{
        phase = 0
        usersCreated = 0
        impersonationReady = $false
        attachmentsReady = $false
        emailsSent = 0
        estimatedSizeGB = 0.0
        newMessagesSent = 0
        repliesSent = 0
        forwardsSent = 0
        lastUserIndex = 0
        startTime = (Get-Date).ToString("o")
    }
}

function Save-State {
    param($State)
    $State | ConvertTo-Json -Depth 10 | Set-Content $StateFile -Encoding UTF8
}

function Get-Threads {
    if (Test-Path $ThreadsFile) {
        return Get-Content $ThreadsFile -Raw | ConvertFrom-Json
    }
    return @{ messages = @() }
}

function Save-Threads {
    param($Threads)
    $Threads | ConvertTo-Json -Depth 10 -Compress | Set-Content $ThreadsFile -Encoding UTF8
}

function New-RandomPassword {
    param([int]$Length = 12)
    $upper = "ABCDEFGHJKLMNPQRSTUVWXYZ"
    $lower = "abcdefghjkmnpqrstuvwxyz"
    $digits = "23456789"
    $special = "!@#$%&*"
    $all = $upper + $lower + $digits + $special

    # Guarantee at least one of each type
    $pwd = ""
    $pwd += $upper[(Get-Random -Maximum $upper.Length)]
    $pwd += $lower[(Get-Random -Maximum $lower.Length)]
    $pwd += $digits[(Get-Random -Maximum $digits.Length)]
    $pwd += $special[(Get-Random -Maximum $special.Length)]

    for ($i = 4; $i -lt $Length; $i++) {
        $pwd += $all[(Get-Random -Maximum $all.Length)]
    }

    # Shuffle
    $chars = $pwd.ToCharArray()
    for ($i = $chars.Length - 1; $i -gt 0; $i--) {
        $j = Get-Random -Maximum ($i + 1)
        $tmp = $chars[$i]; $chars[$i] = $chars[$j]; $chars[$j] = $tmp
    }
    return -join $chars
}

#####################################################################
# INTERNATIONAL NAMES (for realistic display names)
#####################################################################

$FirstNames = @(
    "James","Maria","Wei","Olga","Ahmed","Yuki","Hans","Priya","Carlos","Fatima",
    "John","Anna","Hiroshi","Elena","Mohammed","Sakura","Klaus","Deepa","Miguel","Aisha",
    "Robert","Sofia","Chen","Natasha","Ali","Keiko","Franz","Ananya","Pedro","Layla",
    "David","Isabella","Jun","Svetlana","Omar","Hana","Wolfgang","Kavya","Andres","Noor",
    "Michael","Lucia","Takeshi","Irina","Hassan","Rin","Dieter","Meera","Juan","Zara",
    "William","Camila","Liang","Tatiana","Ibrahim","Yui","Markus","Riya","Diego","Amira",
    "Thomas","Valentina","Sato","Anastasia","Yusuf","Aoi","Stefan","Nisha","Luis","Hala",
    "Daniel","Gabriela","Kenji","Marina","Khalid","Mio","Tobias","Sita","Alejandro","Dina",
    "Richard","Paula","Ryu","Vera","Tariq","Emi","Lukas","Pooja","Fernando","Sara",
    "Joseph","Andrea","Kaito","Darya","Faisal","Nana","Felix","Aditi","Mateo","Yasmin",
    "Andrew","Laura","Shota","Alina","Mansoor","Saki","Jens","Tara","Rafael","Leila",
    "Alex","Carmen","Daichi","Polina","Rashid","Miyu","Peter","Divya","Pablo","Amina",
    "Ryan","Martina","Haruto","Kristina","Samir","Yuna","Uwe","Arya","Emilio","Rana",
    "Nathan","Jessica","Sota","Yulia","Hamza","Akari","Bernd","Neha","Sergio","Huda",
    "Brian","Nicole","Ren","Masha","Nabil","Koharu","Ralf","Isha","Victor","Dalia"
)

$LastNames = @(
    "Smith","Garcia","Wang","Ivanov","Al-Said","Tanaka","Mueller","Sharma","Silva","Hassan",
    "Johnson","Martinez","Li","Petrov","Ahmed","Yamamoto","Schmidt","Patel","Santos","Ibrahim",
    "Williams","Lopez","Zhang","Sokolov","Khalil","Suzuki","Schneider","Gupta","Oliveira","Ali",
    "Brown","Hernandez","Liu","Kuznetsov","Omar","Takahashi","Fischer","Singh","Pereira","Mahmoud",
    "Jones","Gonzalez","Chen","Popov","Malik","Watanabe","Weber","Kumar","Costa","Yusuf",
    "Davis","Rodriguez","Yang","Volkov","Rahman","Ito","Meyer","Reddy","Ferreira","Mustafa",
    "Miller","Perez","Huang","Morozov","Hasan","Nakamura","Wagner","Joshi","Almeida","Rashid",
    "Wilson","Sanchez","Wu","Novikov","Karim","Kobayashi","Becker","Verma","Souza","Saleh",
    "Moore","Ramirez","Zhou","Kozlov","Hussain","Kato","Schulz","Rao","Lima","Hamid",
    "Taylor","Torres","Xu","Lebedev","Farooq","Yoshida","Hoffmann","Nair","Rocha","Osman",
    "Anderson","Flores","Sun","Sorokin","Akhtar","Yamada","Bauer","Pillai","Ribeiro","Abbas",
    "White","Rivera","Ma","Pavlov","Siddiqui","Sasaki","Koch","Iyer","Martins","Nasser",
    "Harris","Gomez","Zhu","Semenov","Qureshi","Yamaguchi","Richter","Shah","Barbosa","Farid",
    "Clark","Diaz","Gao","Egorov","Chaudhry","Matsumoto","Klein","Mishra","Araujo","Kareem",
    "Lewis","Cruz","Lin","Fedorov","Raza","Inoue","Kraus","Tiwari","Cardoso","Saeed"
)

#####################################################################
# MULTI-LANGUAGE TEXT SNIPPETS (50-100 chars each)
#####################################################################

$TextSnippets = @(
    # English
    "Please find the attached report for your review and consideration.",
    "The quarterly results exceeded expectations by a significant margin.",
    "Could you schedule a meeting for next week to discuss the project?",
    "Thank you for your prompt response regarding the recent inquiry.",
    "I have updated the document based on the feedback received today.",

    # Russian
    "Пожалуйста, ознакомьтесь с прикреплённым отчётом для рассмотрения.",
    "Квартальные результаты значительно превысили ожидания руководства.",
    "Не могли бы вы назначить встречу на следующей неделе для обсуждения?",
    "Благодарю за оперативный ответ по поводу недавнего запроса.",
    "Я обновил документ на основе полученных сегодня замечаний и правок.",

    # Spanish
    "Por favor, revise el informe adjunto para su debida consideracion.",
    "Los resultados trimestrales superaron las expectativas notablemente.",
    "Podria programar una reunion la proxima semana para discutir esto?",
    "Gracias por su pronta respuesta respecto a la consulta reciente.",
    "He actualizado el documento basandome en los comentarios recibidos.",

    # French
    "Veuillez trouver le rapport ci-joint pour votre examen attentif.",
    "Les resultats trimestriels ont largement depasse les previsions.",
    "Pourriez-vous organiser une reunion la semaine prochaine pour cela?",
    "Merci pour votre reponse rapide concernant la demande recente.",
    "Le document a ete mis a jour selon les retours recus aujourd'hui.",

    # German
    "Bitte sehen Sie den beigefuegten Bericht zur Pruefung und Bewertung.",
    "Die Quartalsergebnisse uebertrafen die Erwartungen deutlich.",
    "Koennten Sie naechste Woche ein Meeting dafuer einplanen bitte?",
    "Vielen Dank fuer Ihre schnelle Antwort bezueglich der Anfrage.",
    "Ich habe das Dokument basierend auf dem Feedback aktualisiert.",

    # Chinese
    "请查阅附件中的报告并提供您的审核意见和建议。",
    "季度业绩大幅超出预期目标，取得了显著成效。",
    "能否安排下周开会讨论该项目的进展和计划？",
    "感谢您对最近查询事项的及时回复和处理反馈。",
    "我已根据今天收到的反馈意见更新了相关文档。",

    # Japanese
    "添付のレポートをご確認の上、ご検討いただけますようお願いいたします。",
    "四半期の業績は予想を大幅に上回る結果となりました。",
    "来週、プロジェクトについて会議の予定を立てていただけますか？",
    "先日のお問い合わせに対する迅速なご対応ありがとうございます。",
    "本日いただいたフィードバックに基づき、文書を更新いたしました。",

    # Arabic
    "يرجى مراجعة التقرير المرفق للنظر فيه وتقديم ملاحظاتكم.",
    "تجاوزت النتائج الفصلية التوقعات بشكل ملحوظ وكبير جداً.",
    "هل يمكنكم جدولة اجتماع الأسبوع القادم لمناقشة المشروع؟",
    "شكراً لردكم السريع بخصوص الاستفسار الأخير المقدم منا.",
    "لقد قمت بتحديث المستند بناءً على الملاحظات الواردة اليوم.",

    # Portuguese
    "Por favor, verifique o relatorio anexo para sua devida analise.",
    "Os resultados trimestrais superaram significativamente as metas.",
    "Poderia agendar uma reuniao na proxima semana para discutirmos?",
    "Obrigado pela resposta rapida sobre a consulta feita recentemente.",
    "Atualizei o documento com base no feedback recebido hoje pela manha.",

    # Italian
    "Si prega di esaminare il rapporto allegato per la sua valutazione.",
    "I risultati trimestrali hanno superato notevolmente le aspettative.",
    "Potrebbe fissare una riunione la prossima settimana per discuterne?",
    "Grazie per la rapida risposta riguardo alla recente richiesta fatta.",
    "Ho aggiornato il documento in base al feedback ricevuto quest'oggi."
)

$EmailSubjects = @(
    "Quarterly Report Q4 2025", "Meeting Request: Project Update",
    "Re: Budget Approval", "FW: Conference Details", "Action Required: Review Document",
    "Weekly Status Update", "Team Sync — Priorities", "Follow-up: Client Feedback",
    "Updated Proposal Draft", "Infrastructure Maintenance Notice",
    "New Policy Guidelines", "Training Schedule Update", "Vendor Evaluation Results",
    "Project Timeline Revision", "Security Audit Findings", "Product Launch Roadmap",
    "Customer Survey Results", "Office Relocation Plan", "Annual Review Preparation",
    "Partnership Opportunity", "Technical Specification Review", "Compliance Update",
    "Budget Forecast FY2026", "Team Building Event", "Service Level Agreement Draft",
    "Data Migration Status", "Performance Metrics Report", "Risk Assessment Summary",
    "Recruitment Update", "IT Support Ticket Summary",
    "Отчёт за квартал", "Приглашение на совещание", "Обновление проекта",
    "Rapport trimestriel", "Invitation reunion", "Mise a jour du projet",
    "Informe trimestral", "Solicitud de reunion", "Actualizacion del proyecto",
    "Quartalsbericht", "Besprechungseinladung", "Projekt-Update",
    "四半期報告書", "会議の招待", "プロジェクト更新",
    "季度报告", "会议邀请", "项目更新进展",
    "Relatorio trimestral", "Convite para reuniao", "Atualizacao do projeto"
)

#####################################################################
# HTML EMAIL TEMPLATES
#####################################################################

function Get-HtmlSignature {
    param([string]$DisplayName, [string]$Email)
    $titles = @("Senior Analyst","Project Manager","IT Specialist","Team Lead",
                "Consultant","Engineer","Director","Coordinator","Administrator","Architect")
    $depts = @("IT","Finance","Operations","HR","Engineering","Sales","Marketing","Support","Legal","R&D")
    $phones = @("+1-555-","+ 7-495-","+44-20-","+49-30-","+33-1-","+81-3-","+86-10-","+61-2-")

    $title = $titles | Get-Random
    $dept = $depts | Get-Random
    $phone = ($phones | Get-Random) + (Get-Random -Minimum 1000000 -Maximum 9999999)

    return @"
<br/>
<table style="font-family:Arial,sans-serif;font-size:11px;color:#555;border-top:2px solid #4472C4;padding-top:8px;margin-top:15px;">
<tr><td style="font-weight:bold;font-size:13px;color:#333;">$DisplayName</td></tr>
<tr><td style="color:#4472C4;font-style:italic;">$title — $dept Department</td></tr>
<tr><td>Phone: $phone | Email: <a href="mailto:$Email" style="color:#4472C4;">$Email</a></td></tr>
<tr><td style="font-size:10px;color:#999;padding-top:4px;">This email and any attachments are confidential.</td></tr>
</table>
"@
}

function Get-HtmlTemplate_BusinessFormal {
    param([string]$BodyText, [string]$Signature, [string]$RecipientName)
    return @"
<html><head><meta charset="utf-8"></head>
<body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#333;line-height:1.6;">
<p>Dear $RecipientName,</p>
<p>$BodyText</p>
<p>Please do not hesitate to reach out if you have any questions or require further information.</p>
<p>Best regards,</p>
$Signature
</body></html>
"@
}

function Get-HtmlTemplate_Casual {
    param([string]$BodyText, [string]$Signature)
    $colors = @("#e8f4f8","#fff3e0","#e8f5e9","#fce4ec","#f3e5f5")
    $bg = $colors | Get-Random
    return @"
<html><head><meta charset="utf-8"></head>
<body style="font-family:'Segoe UI',Tahoma,sans-serif;font-size:14px;color:#333;">
<div style="background-color:$bg;padding:15px;border-radius:8px;margin:10px 0;">
<p style="margin:0;">Hey!</p>
<p>$BodyText</p>
<p>Cheers!</p>
</div>
$Signature
</body></html>
"@
}

function Get-HtmlTemplate_Report {
    param([string]$BodyText, [string]$Signature)
    # Generate a small mock data table
    $rows = ""
    $metrics = @("Revenue","Costs","Profit","Users","Uptime","Tickets","SLA","Capacity","Latency","Throughput")
    for ($i = 0; $i -lt 5; $i++) {
        $metric = $metrics | Get-Random
        $val = Get-Random -Minimum 50 -Maximum 9999
        $pct = Get-Random -Minimum -20 -Maximum 40
        $pctColor = if ($pct -ge 0) { "#28a745" } else { "#dc3545" }
        $pctSign = if ($pct -ge 0) { "+" } else { "" }
        $rows += "<tr><td style='padding:6px 12px;border-bottom:1px solid #eee;'>$metric</td>"
        $rows += "<td style='padding:6px 12px;border-bottom:1px solid #eee;text-align:right;'>$val</td>"
        $rows += "<td style='padding:6px 12px;border-bottom:1px solid #eee;text-align:right;color:$pctColor;font-weight:bold;'>$pctSign$pct%</td></tr>`n"
    }
    return @"
<html><head><meta charset="utf-8"></head>
<body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#333;line-height:1.6;">
<h2 style="color:#4472C4;border-bottom:2px solid #4472C4;padding-bottom:5px;">Status Report</h2>
<p>$BodyText</p>
<table style="border-collapse:collapse;width:100%;margin:15px 0;font-size:13px;">
<tr style="background-color:#4472C4;color:white;">
<th style="padding:8px 12px;text-align:left;">Metric</th>
<th style="padding:8px 12px;text-align:right;">Value</th>
<th style="padding:8px 12px;text-align:right;">Change</th>
</tr>
$rows
</table>
<p style="font-size:12px;color:#666;">Report generated automatically. Data is for illustration purposes.</p>
$Signature
</body></html>
"@
}

function Get-HtmlTemplate_Newsletter {
    param([string]$BodyText, [string]$Signature, [string]$InlineImageCid)
    $headerImg = ""
    if ($InlineImageCid) {
        $headerImg = "<img src='cid:$InlineImageCid' style='width:100%;max-height:200px;object-fit:cover;border-radius:8px 8px 0 0;' alt='Header'/>"
    }
    return @"
<html><head><meta charset="utf-8"></head>
<body style="font-family:'Segoe UI',Tahoma,sans-serif;font-size:14px;color:#333;background-color:#f5f5f5;padding:20px;">
<div style="max-width:600px;margin:0 auto;background:white;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1);">
$headerImg
<div style="padding:20px;">
<h1 style="color:#4472C4;font-size:20px;margin-top:0;">Weekly Newsletter</h1>
<p>$BodyText</p>
<div style="background:#f0f4f8;padding:12px;border-left:4px solid #4472C4;margin:15px 0;border-radius:0 4px 4px 0;">
<strong>Highlights:</strong>
<ul style="margin:5px 0;padding-left:20px;">
<li>New team members joined this month</li>
<li>Upcoming deadlines and milestones</li>
<li>Updated guidelines available on the portal</li>
</ul>
</div>
</div>
<div style="background:#343a40;color:white;padding:15px 20px;border-radius:0 0 8px 8px;font-size:11px;">
Internal communication — Do not forward externally.
</div>
</div>
$Signature
</body></html>
"@
}

function Get-HtmlTemplate_SimpleReply {
    param([string]$BodyText, [string]$Signature)
    return @"
<html><head><meta charset="utf-8"></head>
<body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#333;line-height:1.6;">
<p>$BodyText</p>
$Signature
</body></html>
"@
}

function Get-RandomHtmlBody {
    param(
        [string]$SenderName,
        [string]$SenderEmail,
        [string]$RecipientName,
        [string]$InlineImageCid
    )

    $text = $TextSnippets | Get-Random
    $sig = Get-HtmlSignature -DisplayName $SenderName -Email $SenderEmail
    $templateChoice = Get-Random -Minimum 1 -Maximum 6

    switch ($templateChoice) {
        1 { return Get-HtmlTemplate_BusinessFormal -BodyText $text -Signature $sig -RecipientName $RecipientName }
        2 { return Get-HtmlTemplate_Casual -BodyText $text -Signature $sig }
        3 { return Get-HtmlTemplate_Report -BodyText $text -Signature $sig }
        4 { return Get-HtmlTemplate_Newsletter -BodyText $text -Signature $sig -InlineImageCid $InlineImageCid }
        5 { return Get-HtmlTemplate_SimpleReply -BodyText $text -Signature $sig }
    }
}

#####################################################################
# BANNER
#####################################################################

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Exchange Mock Data Generator" -ForegroundColor Cyan
Write-Host "  Target: $TargetSizeGB GB | Users: $UserCount | Threads: $Threads" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

$State = Get-State
$StartPhase = if ($SkipToPhase -gt 0) { $SkipToPhase } elseif ($State.phase -gt 0) { $State.phase } else { 1 }

#####################################################################
# PHASE 1: CREATE MOCK USERS
#####################################################################

if ($StartPhase -le 1) {
    Write-Host "=== PHASE 1: Creating $UserCount Mock Users ===" -ForegroundColor Yellow
    Write-Host ""

    # Load Exchange snap-in if needed
    if (-not (Get-Command New-Mailbox -ErrorAction SilentlyContinue)) {
        try {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
            Write-Log "Exchange Management snap-in loaded" "OK"
        } catch {
            Write-Log "Failed to load Exchange snap-in. Run from Exchange Management Shell." "ERROR"
            exit 1
        }
    }

    # Detect domain and pin to a single DC for all operations
    try {
        $ADDomain = Get-ADDomain
        $DomainDN = $ADDomain.DistinguishedName
        $DomainFQDN = $ADDomain.DNSRoot
        # Use PDC emulator — guaranteed single writable DC, no replication lag
        $DC = $ADDomain.PDCEmulator
        Write-Log "Domain detected: $DomainFQDN ($DomainDN)" "OK"
        Write-Log "Pinned to DC: $DC (PDC Emulator)" "OK"
    } catch {
        Write-Log "Failed to detect AD domain: $_" "ERROR"
        exit 1
    }

    # Store DC for other phases
    $script:TargetDC = $DC

    # Create OU if not exists (pinned to DC)
    $OUPath = "OU=$MockUsersOU,$DomainDN"
    try {
        Get-ADOrganizationalUnit -Identity $OUPath -Server $DC -ErrorAction Stop | Out-Null
        Write-Log "OU '$MockUsersOU' already exists" "OK"
    } catch {
        try {
            New-ADOrganizationalUnit -Name $MockUsersOU -Path $DomainDN -Server $DC -ProtectedFromAccidentalDeletion $false
            Write-Log "OU '$MockUsersOU' created on $DC" "OK"

            # Verify OU is visible on the same DC before proceeding
            $retries = 0
            while ($retries -lt 15) {
                try {
                    Get-ADOrganizationalUnit -Identity $OUPath -Server $DC -ErrorAction Stop | Out-Null
                    Write-Log "OU verified on $DC" "OK"
                    break
                } catch {
                    $retries++
                    Write-Log "  Waiting for OU to be available on $DC (attempt $retries/15)..." "WARN"
                    Start-Sleep -Seconds 2
                }
            }
            if ($retries -ge 15) {
                Write-Log "OU not available after 30 seconds. Exiting." "ERROR"
                exit 1
            }
        } catch {
            Write-Log "Failed to create OU: $_" "ERROR"
            exit 1
        }
    }

    # Detect mailbox database
    $MDB = (Get-MailboxDatabase -DomainController $DC | Select-Object -First 1).Name
    if (-not $MDB) {
        Write-Log "No mailbox database found" "ERROR"
        exit 1
    }
    Write-Log "Using mailbox database: $MDB"

    # Prepare credentials CSV
    $CredsData = @()
    $Created = 0
    $Skipped = 0

    for ($i = 1; $i -le $UserCount; $i++) {
        $num = $i.ToString("D3")
        $alias = "$UserPrefix$num"
        $upn = "$alias@$DomainFQDN"

        # Pick random international name
        $firstName = $FirstNames[(($i - 1) % $FirstNames.Count)]
        $lastName = $LastNames[(($i - 1) % $LastNames.Count)]
        $displayName = "$firstName $lastName"

        # Check if already exists (pinned to DC)
        $existing = Get-Mailbox -Identity $alias -DomainController $DC -ErrorAction SilentlyContinue
        if ($existing) {
            $Skipped++
            # Still record in CSV if not already there
            $password = "***existing***"
            $CredsData += [PSCustomObject]@{
                Number = $i
                Alias = $alias
                UPN = $upn
                DisplayName = $displayName
                Password = $password
                SamAccountName = $alias
            }
            continue
        }

        $password = New-RandomPassword -Length $PasswordLength
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force

        try {
            # Use alias as -Name (AD CN) to guarantee uniqueness;
            # international name goes into -DisplayName only
            New-Mailbox -Name $alias `
                        -Alias $alias `
                        -UserPrincipalName $upn `
                        -SamAccountName $alias `
                        -FirstName $firstName `
                        -LastName $lastName `
                        -DisplayName $displayName `
                        -Database $MDB `
                        -OrganizationalUnit $OUPath `
                        -Password $securePassword `
                        -ResetPasswordOnNextLogon:$false `
                        -DomainController $DC `
                        -ErrorAction Stop | Out-Null

            $Created++
            Write-Log "  [$i/$UserCount] Created: $upn ($displayName)" "OK"
        } catch {
            Write-Log "  [$i/$UserCount] Failed: $upn — $_" "ERROR"
            $password = "***FAILED***"
        }

        $CredsData += [PSCustomObject]@{
            Number = $i
            Alias = $alias
            UPN = $upn
            DisplayName = $displayName
            Password = $password
            SamAccountName = $alias
        }

        # Small delay to not overload AD
        if ($i % 10 -eq 0) {
            Start-Sleep -Milliseconds 500
        }
    }

    # Export credentials
    $CredsData | Export-Csv -Path $CredsFile -NoTypeInformation -Encoding UTF8
    Write-Log "Credentials exported to: $CredsFile" "OK"
    Write-Log "Created: $Created | Skipped (existing): $Skipped"

    $State.usersCreated = $UserCount
    $State.phase = 2
    Save-State $State

    Write-Host ""
    Write-Host "--- Phase 1 Complete ---" -ForegroundColor Green
    Write-Host ""
}

#####################################################################
# PHASE 2: VALIDATE SMTP RELAY
#####################################################################

if ($StartPhase -le 2) {
    Write-Host "=== PHASE 2: Configure Authenticated SMTP ===" -ForegroundColor Yellow
    Write-Host ""

    # Ensure DC is set (in case we're resuming from Phase 2)
    if (-not $script:TargetDC) {
        $script:TargetDC = (Get-ADDomain).PDCEmulator
        Write-Log "Pinned to DC: $($script:TargetDC) (PDC Emulator)" "OK"
    }
    $DC = $script:TargetDC

    # Detect domain for email addresses
    $DomainFQDN = if ($DomainFQDN) { $DomainFQDN } else { (Get-ADDomain -Server $DC).DNSRoot }

    # ---- Select SMTP server ----
    if ([string]::IsNullOrEmpty($SmtpServer)) {
        Write-Log "Detecting Exchange servers..."
        $ExServers = @()
        try {
            # Exchange 2016/2019 — all Mailbox role servers have CAS
            $ExServers = @(Get-ExchangeServer | Where-Object {
                $_.ServerRole -match "Mailbox" -and $_.AdminDisplayVersion -match "15\."
            } | Sort-Object Name)
        } catch {
            try {
                # Fallback: Get-ClientAccessServer (older Exchange versions)
                $ExServers = @(Get-ClientAccessServer | Sort-Object Name)
            } catch {
                Write-Log "Cannot enumerate Exchange servers: $_" "WARN"
            }
        }

        if ($ExServers.Count -eq 0) {
            # Last resort: use local server FQDN
            $SmtpServer = [System.Net.Dns]::GetHostEntry([System.Net.Dns]::GetHostName()).HostName
            Write-Log "No Exchange servers found via cmdlets, using local FQDN: $SmtpServer" "WARN"
        } elseif ($ExServers.Count -eq 1) {
            $SmtpServer = $ExServers[0].Fqdn
            if (-not $SmtpServer) { $SmtpServer = "$($ExServers[0].Name).$DomainFQDN" }
            Write-Log "Single Exchange server detected: $SmtpServer" "OK"
        } else {
            Write-Host ""
            Write-Host "  Available Exchange servers:" -ForegroundColor Cyan
            for ($si = 0; $si -lt $ExServers.Count; $si++) {
                $srvFqdn = $ExServers[$si].Fqdn
                if (-not $srvFqdn) { $srvFqdn = "$($ExServers[$si].Name).$DomainFQDN" }
                $roles = $ExServers[$si].ServerRole
                Write-Host "    [$($si + 1)] $srvFqdn  ($roles)" -ForegroundColor White
            }
            Write-Host ""
            $choice = Read-Host "  Select server number (1-$($ExServers.Count))"
            $choiceIdx = [int]$choice - 1
            if ($choiceIdx -lt 0 -or $choiceIdx -ge $ExServers.Count) {
                Write-Log "Invalid choice, using first server" "WARN"
                $choiceIdx = 0
            }
            $SmtpServer = $ExServers[$choiceIdx].Fqdn
            if (-not $SmtpServer) { $SmtpServer = "$($ExServers[$choiceIdx].Name).$DomainFQDN" }
            Write-Log "Selected SMTP server: $SmtpServer" "OK"
        }
    } else {
        Write-Log "Using specified SMTP server: $SmtpServer"
    }

    # Store in state for Phase 4
    $State.smtpServer = $SmtpServer
    $State.smtpPort = $SmtpPort

    # ---- Test SMTP connectivity ----
    Write-Log "Testing SMTP connectivity: ${SmtpServer}:${SmtpPort}..."
    try {
        $tcpTest = New-Object System.Net.Sockets.TcpClient
        $tcpTest.Connect($SmtpServer, $SmtpPort)
        $tcpTest.Close()
        Write-Log "SMTP port $SmtpPort is open on $SmtpServer" "OK"
    } catch {
        Write-Log "SMTP port $SmtpPort is not accessible on ${SmtpServer}: $_" "ERROR"
        Write-Log "Ensure SMTP is enabled on that server. Check:" "ERROR"
        Write-Log "  Get-ReceiveConnector | FL Name,Bindings,AuthMechanism" "ERROR"
        exit 1
    }

    # ---- Test authenticated send ----
    # Load credentials for test user
    if (-not (Test-Path $CredsFile)) {
        Write-Log "Credentials file not found: $CredsFile — run Phase 1 first" "ERROR"
        exit 1
    }
    $TestUsers = Import-Csv $CredsFile | Where-Object { $_.Password -ne "***FAILED***" -and $_.Password -ne "***existing***" }
    if ($TestUsers.Count -lt 2) {
        Write-Log "Need at least 2 users with known passwords in $CredsFile" "ERROR"
        exit 1
    }

    $testSender = $TestUsers[0]
    $testRecipient = $TestUsers[1]

    Write-Log "Sending authenticated test email: $($testSender.UPN) -> $($testRecipient.UPN)..."
    try {
        $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
        $smtp.EnableSsl = $true
        $smtp.Credentials = New-Object System.Net.NetworkCredential($testSender.UPN, $testSender.Password)

        $testMsg = New-Object System.Net.Mail.MailMessage(
            $testSender.UPN, $testRecipient.UPN,
            "SMTP Auth Test", "Test message from mock data generator (authenticated SMTP on port $SmtpPort).")
        $smtp.Send($testMsg)
        $testMsg.Dispose()
        $smtp.Dispose()
        Write-Log "Authenticated SMTP test PASSED: $($testSender.UPN) -> $($testRecipient.UPN)" "OK"
    } catch {
        Write-Log "Authenticated SMTP test FAILED: $_" "ERROR"
        Write-Log "" "ERROR"
        Write-Log "Troubleshooting:" "WARN"
        Write-Log "  1. Check Receive Connector auth: Get-ReceiveConnector | FL Name,Bindings,AuthMechanism,PermissionGroups" "WARN"
        Write-Log "  2. Ensure users have SMTP send enabled: Get-CASMailbox $($testSender.Alias) | FL SmtpClientAuthenticationDisabled" "WARN"
        Write-Log "  3. Try port 587 instead: .\Generate-ExchangeMockData.ps1 -SmtpPort 587 -SkipToPhase 2" "WARN"
        exit 1
    }

    # ---- Raise throttling limits for bulk sending ----
    Write-Log "Configuring Exchange throttling policy for bulk sending..."
    try {
        # Create or update a throttling policy with unlimited rate for mock users
        $policyName = "MockDataBulkSend"
        $existingPolicy = Get-ThrottlingPolicy -Identity $policyName -ErrorAction SilentlyContinue
        if ($existingPolicy) {
            Set-ThrottlingPolicy -Identity $policyName `
                -MessageRateLimit Unlimited `
                -RecipientRateLimit Unlimited `
                -ErrorAction Stop
            Write-Log "  Updated throttling policy '$policyName': MessageRateLimit=Unlimited, RecipientRateLimit=Unlimited" "OK"
        } else {
            New-ThrottlingPolicy -Name $policyName `
                -MessageRateLimit Unlimited `
                -RecipientRateLimit Unlimited `
                -ThrottlingPolicyScope Regular `
                -ErrorAction Stop
            Write-Log "  Created throttling policy '$policyName': MessageRateLimit=Unlimited, RecipientRateLimit=Unlimited" "OK"
        }

        # Apply policy to all mock users
        $applied = 0
        for ($pi = 1; $pi -le $UserCount; $pi++) {
            $alias = "${UserPrefix}$($pi.ToString('D3'))"
            try {
                Set-Mailbox -Identity $alias -ThrottlingPolicy $policyName -DomainController $DC -ErrorAction Stop
                $applied++
            } catch { }
            if ($pi % 50 -eq 0) { Write-Log "  Applied to $applied/$pi users..." }
        }
        Write-Log "  Applied '$policyName' to $applied mailboxes" "OK"
    } catch {
        Write-Log "  Failed to set throttling policy: $_" "WARN"
        Write-Log "  You can set it manually:" "WARN"
        Write-Log "    New-ThrottlingPolicy -Name MockDataBulkSend -MessageRateLimit Unlimited -RecipientRateLimit Unlimited -ThrottlingPolicyScope Regular" "WARN"
        Write-Log "    Get-Mailbox -OrganizationalUnit $MockUsersOU | Set-Mailbox -ThrottlingPolicy MockDataBulkSend" "WARN"
    }

    # Also raise Receive Connector message rate limit
    try {
        $connectors = Get-ReceiveConnector | Where-Object { $_.Bindings -match ":465|:587" }
        foreach ($rc in $connectors) {
            Set-ReceiveConnector -Identity $rc.Identity -MessageRateLimit unlimited -ErrorAction Stop
            Write-Log "  Set MessageRateLimit=unlimited on '$($rc.Identity)'" "OK"
        }
    } catch {
        Write-Log "  Could not adjust Receive Connector rate limit: $_" "WARN"
        Write-Log "  You can do it manually: Get-ReceiveConnector | Set-ReceiveConnector -MessageRateLimit unlimited" "WARN"
    }

    # Raise mailbox database delivery limits (fixes 4.3.2 errors)
    Write-Log "Raising Transport/Mailbox delivery throttling (4.3.2 prevention)..."
    $targetServer = $SmtpServer.Split('.')[0]  # short hostname

    # TransportService — MaxConcurrentMailboxDeliveries / MaxConcurrentMailboxSubmissions
    try {
        Set-TransportService -Identity $targetServer `
            -MaxConcurrentMailboxDeliveries 100 `
            -MaxConcurrentMailboxSubmissions 100 `
            -ErrorAction Stop
        Write-Log "  Set-TransportService $targetServer : MaxConcurrentMailboxDeliveries=100, MaxConcurrentMailboxSubmissions=100" "OK"
    } catch {
        Write-Log "  Could not set TransportService limits on '$targetServer': $_" "WARN"
    }

    # MailboxTransportDeliveryService — MaxConcurrentMailboxDeliveries
    try {
        Set-MailboxTransportService -Identity $targetServer `
            -MaxConcurrentMailboxDeliveries 100 `
            -ErrorAction Stop
        Write-Log "  Set-MailboxTransportService $targetServer : MaxConcurrentMailboxDeliveries=100" "OK"
    } catch {
        Write-Log "  Could not set MailboxTransportService limits: $_" "WARN"
    }

    # MailboxTransportSubmissionService — MaxConcurrentMailboxSubmissions
    try {
        Set-MailboxTransportService -Identity $targetServer `
            -MaxConcurrentMailboxSubmissions 100 `
            -ErrorAction Stop
        Write-Log "  Set-MailboxTransportService $targetServer : MaxConcurrentMailboxSubmissions=100" "OK"
    } catch {
        # May not exist separately in all Exchange versions
    }

    # If there is a second Exchange server, apply same settings
    try {
        $allExServers = @(Get-ExchangeServer | Where-Object { $_.ServerRole -match "Mailbox" } | Select-Object -ExpandProperty Name)
        foreach ($srv in $allExServers) {
            if ($srv -eq $targetServer) { continue }
            try {
                Set-TransportService -Identity $srv -MaxConcurrentMailboxDeliveries 100 -MaxConcurrentMailboxSubmissions 100 -ErrorAction Stop
                Set-MailboxTransportService -Identity $srv -MaxConcurrentMailboxDeliveries 100 -ErrorAction Stop
                Write-Log "  Also raised limits on '$srv'" "OK"
            } catch {
                Write-Log "  Could not set limits on '$srv': $_" "WARN"
            }
        }
    } catch { }

    # Raise per-database replication/delivery limits via Set-MailboxDatabase
    try {
        $databases = Get-MailboxDatabase -DomainController $DC
        foreach ($db in $databases) {
            Set-MailboxDatabase -Identity $db.Name `
                -MailboxRetention 0.00:00:00 `
                -ErrorAction SilentlyContinue
            Write-Log "  Database '$($db.Name)': retention set to minimum" "OK"
        }
    } catch {
        Write-Log "  Could not adjust database settings: $_" "WARN"
    }

    Write-Log "Transport throttling configuration complete" "OK"

    $State.smtpReady = $true
    $State.phase = 3
    Save-State $State

    Write-Host ""
    Write-Host "--- Phase 2 Complete ---" -ForegroundColor Green
    Write-Host ""
}

#####################################################################
# PHASE 3: DOWNLOAD / GENERATE SAMPLE ATTACHMENTS
#####################################################################

if ($StartPhase -le 3) {
    Write-Host "=== PHASE 3: Preparing Sample Attachments ===" -ForegroundColor Yellow
    Write-Host ""

    # Ensure directories exist
    @($JpgDir, $TxtDir, $RtfDir) | ForEach-Object {
        if (-not (Test-Path $_)) { New-Item -ItemType Directory -Path $_ -Force | Out-Null }
    }

    # Enable all TLS versions and bypass cert validation for outbound downloads
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
    } catch {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }

    # --- Download JPG images from picsum.photos ---
    Write-Log "Downloading JPG images..."

    $jpgSpecs = @(
        @{ prefix = "small"; width = 640; height = 480; count = 20 },
        @{ prefix = "medium"; width = 1920; height = 1080; count = 20 },
        @{ prefix = "large"; width = 3840; height = 2160; count = 10 }
    )

    # Try multiple image sources in case one is blocked
    $imageSources = @(
        { param($w,$h) "https://picsum.photos/$w/$h" },
        { param($w,$h) "https://placehold.co/${w}x${h}.jpg" },
        { param($w,$h) "https://dummyimage.com/${w}x${h}/$('{0:x6}' -f (Get-Random -Max 0xFFFFFF))/fff.jpg" }
    )

    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

    foreach ($spec in $jpgSpecs) {
        for ($i = 1; $i -le $spec.count; $i++) {
            $filename = "$($spec.prefix)_$($i.ToString('D2')).jpg"
            $filepath = Join-Path $JpgDir $filename

            if (Test-Path $filepath) {
                continue  # Skip already downloaded
            }

            $downloaded = $false
            foreach ($srcFunc in $imageSources) {
                try {
                    $url = & $srcFunc $spec.width $spec.height
                    # Use WebClient — more compatible with older TLS stacks than Invoke-WebRequest
                    $wc = New-Object System.Net.WebClient
                    $wc.Headers.Add("User-Agent", "Mozilla/5.0")
                    $wc.DownloadFile($url, $filepath)
                    $wc.Dispose()

                    if ((Test-Path $filepath) -and (Get-Item $filepath).Length -gt 100) {
                        $size = [math]::Round((Get-Item $filepath).Length / 1KB)
                        Write-Log "  Downloaded: $filename (${size}KB)" "OK"
                        $downloaded = $true
                        break
                    }
                } catch {
                    # Try next source
                }
            }

            if (-not $downloaded) {
                # Generate a varied fallback JPG with gradient + shapes + text
                try {
                    $bmp = New-Object System.Drawing.Bitmap($spec.width, $spec.height)
                    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
                    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

                    # Gradient background
                    $c1 = [System.Drawing.Color]::FromArgb((Get-Random -Max 256),(Get-Random -Max 256),(Get-Random -Max 256))
                    $c2 = [System.Drawing.Color]::FromArgb((Get-Random -Max 256),(Get-Random -Max 256),(Get-Random -Max 256))
                    $gradBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
                        (New-Object System.Drawing.Point(0,0)),
                        (New-Object System.Drawing.Point($spec.width, $spec.height)),
                        $c1, $c2
                    )
                    $graphics.FillRectangle($gradBrush, 0, 0, $spec.width, $spec.height)

                    # Random shapes
                    for ($s = 0; $s -lt 8; $s++) {
                        $sc = [System.Drawing.Color]::FromArgb(80, (Get-Random -Max 256),(Get-Random -Max 256),(Get-Random -Max 256))
                        $sBrush = New-Object System.Drawing.SolidBrush($sc)
                        $sx = Get-Random -Maximum $spec.width
                        $sy = Get-Random -Maximum $spec.height
                        $sw = Get-Random -Minimum 50 -Maximum ([math]::Max(100, $spec.width / 3))
                        $sh = Get-Random -Minimum 50 -Maximum ([math]::Max(100, $spec.height / 3))
                        if ((Get-Random -Max 2) -eq 0) {
                            $graphics.FillEllipse($sBrush, $sx, $sy, $sw, $sh)
                        } else {
                            $graphics.FillRectangle($sBrush, $sx, $sy, $sw, $sh)
                        }
                        $sBrush.Dispose()
                    }

                    # Text overlay
                    $font = New-Object System.Drawing.Font("Arial", [math]::Max(20, $spec.width / 25))
                    $textBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(200,255,255,255))
                    $graphics.DrawString("Mock Image $($spec.prefix) #$i", $font, $textBrush, 20, 20)
                    $font.Dispose(); $textBrush.Dispose(); $gradBrush.Dispose()

                    $bmp.Save($filepath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
                    $graphics.Dispose(); $bmp.Dispose()
                    $size = [math]::Round((Get-Item $filepath).Length / 1KB)
                    Write-Log "  Generated: $filename (${size}KB)" "OK"
                } catch {
                    Write-Log "  Failed to generate $filename : $_" "WARN"
                }
            }

            Start-Sleep -Milliseconds 200  # Be polite to picsum
        }
    }

    # --- Generate TXT files with multi-language content ---
    Write-Log "Generating TXT files with multi-language content..."

    for ($i = 1; $i -le 30; $i++) {
        $filename = "sample_$($i.ToString('D2')).txt"
        $filepath = Join-Path $TxtDir $filename

        if (Test-Path $filepath) { continue }

        # Build content by repeating random snippets to target size
        $targetKB = Get-Random -Minimum 10 -Maximum 500
        $content = New-Object System.Text.StringBuilder

        while ($content.Length -lt ($targetKB * 1024)) {
            $snippet = $TextSnippets | Get-Random
            [void]$content.AppendLine($snippet)
            [void]$content.AppendLine("")
        }

        [System.IO.File]::WriteAllText($filepath, $content.ToString(), [System.Text.Encoding]::UTF8)
        $size = [math]::Round((Get-Item $filepath).Length / 1KB)
        Write-Log "  Generated: $filename (${size}KB)" "OK"
    }

    # --- Generate RTF files with formatted text ---
    Write-Log "Generating RTF files with formatted content..."

    $rtfColors = @(
        "\red255\green0\blue0;",    # Red
        "\red0\green0\blue255;",    # Blue
        "\red0\green128\blue0;",    # Green
        "\red128\green0\blue128;",  # Purple
        "\red255\green128\blue0;"   # Orange
    )

    for ($i = 1; $i -le 20; $i++) {
        $filename = "document_$($i.ToString('D2')).rtf"
        $filepath = Join-Path $RtfDir $filename

        if (Test-Path $filepath) { continue }

        $targetKB = Get-Random -Minimum 50 -Maximum 1000
        $rtf = New-Object System.Text.StringBuilder

        [void]$rtf.Append("{\rtf1\ansi\deff0")
        [void]$rtf.Append("{\fonttbl{\f0\fswiss Calibri;}{\f1\fmodern Courier New;}{\f2\froman Times New Roman;}}")
        [void]$rtf.Append("{\colortbl ;$($rtfColors -join '')}")
        [void]$rtf.AppendLine("\viewkind4\uc1")

        while ($rtf.Length -lt ($targetKB * 1024)) {
            $snippet = $TextSnippets | Get-Random
            # Escape RTF special chars (basic)
            $escaped = $snippet -replace '\\','\\\\' -replace '\{','\\{' -replace '\}','\\}'
            # For non-ASCII, use Unicode escapes
            $rtfText = New-Object System.Text.StringBuilder
            foreach ($ch in $escaped.ToCharArray()) {
                if ([int]$ch -gt 127) {
                    [void]$rtfText.Append("\u$([int]$ch)?")
                } else {
                    [void]$rtfText.Append($ch)
                }
            }

            $fontIdx = Get-Random -Minimum 0 -Maximum 3
            $colorIdx = Get-Random -Minimum 1 -Maximum 6
            $fontSize = @(22, 24, 28, 32, 36) | Get-Random

            [void]$rtf.Append("\pard\f$fontIdx\fs$fontSize\cf$colorIdx ")

            # Random formatting
            $fmt = Get-Random -Minimum 0 -Maximum 4
            switch ($fmt) {
                0 { [void]$rtf.Append("\b $rtfText \b0") }
                1 { [void]$rtf.Append("\i $rtfText \i0") }
                2 { [void]$rtf.Append("\b\i $rtfText \i0\b0") }
                3 { [void]$rtf.Append("$rtfText") }
            }
            [void]$rtf.AppendLine("\par\par")
        }

        [void]$rtf.Append("}")
        [System.IO.File]::WriteAllText($filepath, $rtf.ToString(), [System.Text.Encoding]::ASCII)
        $size = [math]::Round((Get-Item $filepath).Length / 1KB)
        Write-Log "  Generated: $filename (${size}KB)" "OK"
    }

    # Summary
    $jpgCount = (Get-ChildItem $JpgDir -Filter "*.jpg" -ErrorAction SilentlyContinue).Count
    $txtCount = (Get-ChildItem $TxtDir -Filter "*.txt" -ErrorAction SilentlyContinue).Count
    $rtfCount = (Get-ChildItem $RtfDir -Filter "*.rtf" -ErrorAction SilentlyContinue).Count
    Write-Log "Attachments ready: $jpgCount JPGs, $txtCount TXTs, $rtfCount RTFs" "OK"

    $State.attachmentsReady = $true
    $State.phase = 4
    Save-State $State

    Write-Host ""
    Write-Host "--- Phase 3 Complete ---" -ForegroundColor Green
    Write-Host ""
}

#####################################################################
# PHASE 4: SEND EMAILS
#####################################################################

if ($StartPhase -le 4) {
    Write-Host "=== PHASE 4: Sending Emails (~${TargetSizeGB}GB target) ===" -ForegroundColor Yellow
    Write-Host ""

    # Reload state
    $State = Get-State

    # Read credentials
    if (-not (Test-Path $CredsFile)) {
        Write-Log "Credentials file not found: $CredsFile" "ERROR"
        exit 1
    }
    # Filter: only users with known passwords (skip existing/failed)
    $Users = Import-Csv $CredsFile | Where-Object { $_.Password -ne "***FAILED***" -and $_.Password -ne "***existing***" }
    $UserCount = $Users.Count
    Write-Log "Loaded $UserCount users with credentials"

    # Build a hashtable for fast credential lookup by UPN
    $UserCredMap = @{}
    foreach ($u in $Users) { $UserCredMap[$u.UPN] = $u }

    # Load attachments list
    $JpgFiles = @(Get-ChildItem $JpgDir -Filter "*.jpg" -ErrorAction SilentlyContinue)
    $TxtFiles = @(Get-ChildItem $TxtDir -Filter "*.txt" -ErrorAction SilentlyContinue)
    $RtfFiles = @(Get-ChildItem $RtfDir -Filter "*.rtf" -ErrorAction SilentlyContinue)
    $AllAttachments = @($JpgFiles) + @($TxtFiles) + @($RtfFiles)

    if ($AllAttachments.Count -eq 0) {
        Write-Log "No attachments found — run Phase 3 first" "ERROR"
        exit 1
    }
    Write-Log "Attachment pool: $($AllAttachments.Count) files"

    # SMTP setup — use server and port from Phase 2 (saved in state) or parameters
    if (-not $script:TargetDC) { $script:TargetDC = (Get-ADDomain).PDCEmulator }
    $DC = $script:TargetDC
    $DomainFQDN = if ($DomainFQDN) { $DomainFQDN } else { (Get-ADDomain -Server $DC).DNSRoot }

    # Restore SMTP settings from state (set by Phase 2) or use parameters
    if ([string]::IsNullOrEmpty($SmtpServer) -and $State.smtpServer) { $SmtpServer = $State.smtpServer }
    if ($SmtpPort -eq 465 -and $State.smtpPort) { $SmtpPort = [int]$State.smtpPort }
    if ([string]::IsNullOrEmpty($SmtpServer)) {
        $SmtpServer = [System.Net.Dns]::GetHostEntry([System.Net.Dns]::GetHostName()).HostName
        Write-Log "No SMTP server in state, using local FQDN: $SmtpServer" "WARN"
    }

    # Calculate targets
    $TargetSizeBytes = [int64]$TargetSizeGB * 1073741824
    $AvgEmailSizeBytes = 1572864  # ~1.5MB average including copies
    $TotalSendsNeeded = [math]::Ceiling($TargetSizeBytes / $AvgEmailSizeBytes)

    # Distribution
    $NewMessageTarget = [math]::Ceiling($TotalSendsNeeded * 0.50)
    $ReplyTarget       = [math]::Ceiling($TotalSendsNeeded * 0.30)
    $ForwardTarget     = [math]::Ceiling($TotalSendsNeeded * 0.20)

    Write-Log "Target: $TotalSendsNeeded total send operations"
    Write-Log "  New: $NewMessageTarget | Reply: $ReplyTarget | Forward: $ForwardTarget"
    Write-Log "  SMTP: $SmtpServer`:$SmtpPort (SSL + Auth)"

    # Resume tracking
    $emailsSent = [int]$State.emailsSent
    $newSent    = [int]$State.newMessagesSent
    $replySent  = [int]$State.repliesSent
    $fwdSent    = [int]$State.forwardsSent
    $estimatedSize = [double]$State.estimatedSizeGB

    # Thread tracking — store Message-IDs for replies/forwards
    $ThreadData = Get-Threads
    if (-not $ThreadData.messages) {
        $ThreadData = @{ messages = [System.Collections.ArrayList]@() }
    }
    if ($ThreadData.messages -isnot [System.Collections.ArrayList]) {
        $msgList = [System.Collections.ArrayList]@()
        if ($ThreadData.messages) { foreach ($m in $ThreadData.messages) { [void]$msgList.Add($m) } }
        $ThreadData.messages = $msgList
    }

    $startTime = Get-Date

    # Helper: pick random recipients (not the sender)
    function Get-RandomRecipients {
        param([string]$SenderUPN, [int]$Count = 1)
        $others = $Users | Where-Object { $_.UPN -ne $SenderUPN }
        return @($others | Get-Random -Count ([math]::Min($Count, $others.Count)))
    }

    # Helper: select attachment paths based on distribution
    function Get-RandomAttachmentPaths {
        $roll = Get-Random -Minimum 1 -Maximum 101
        if ($roll -le 40) { return @() }
        elseif ($roll -le 70) {
            $small = $AllAttachments | Where-Object { $_.Length -lt 512KB } | Get-Random -Count 1 -ErrorAction SilentlyContinue
            if ($small) { return @($small.FullName) }
            return @(($AllAttachments | Get-Random -Count 1).FullName)
        } elseif ($roll -le 90) {
            $medium = $AllAttachments | Where-Object { $_.Length -ge 512KB -and $_.Length -lt 5MB } | Get-Random -Count 1 -ErrorAction SilentlyContinue
            if ($medium) { return @($medium.FullName) }
            return @(($AllAttachments | Get-Random -Count 1).FullName)
        } else {
            $count = Get-Random -Minimum 1 -Maximum 4
            return @($AllAttachments | Get-Random -Count ([math]::Min($count, $AllAttachments.Count)) | ForEach-Object { $_.FullName })
        }
    }

    # Helper: estimate email size from paths
    function Get-EstimatedEmailSize {
        param([string]$Body, [string[]]$AttachPaths)
        $bodySize = [System.Text.Encoding]::UTF8.GetByteCount($Body)
        $attachSize = 0
        foreach ($p in $AttachPaths) {
            if ($p -and (Test-Path $p)) { $attachSize += (Get-Item $p).Length }
        }
        return ($bodySize + ($attachSize * 1.33) + 5120) * 2
    }

    #=================================================================
    # RUNSPACE POOL SETUP
    #=================================================================

    # Self-contained scriptblock executed in each runspace worker
    $SendScriptBlock = {
        param(
            [string]$SmtpServer, [int]$SmtpPort, [string]$DomainFQDN,
            [string]$From, [string]$FromPassword,
            [string[]]$To, [string]$Subject, [string]$HtmlBody,
            [string[]]$AttachmentPaths, [string]$InlineImagePath, [string]$InlineCid,
            [string]$InReplyTo, [string]$References
        )

        try {
            $mail = New-Object System.Net.Mail.MailMessage
            $mail.From = New-Object System.Net.Mail.MailAddress($From)
            foreach ($addr in $To) { $mail.To.Add($addr) }
            $mail.Subject = $Subject
            $mail.IsBodyHtml = $true
            $mail.SubjectEncoding = [System.Text.Encoding]::UTF8
            $mail.BodyEncoding = [System.Text.Encoding]::UTF8

            $messageId = "<$([guid]::NewGuid().ToString('N'))@$DomainFQDN>"
            $mail.Headers.Add("Message-ID", $messageId)

            if ($InReplyTo) { $mail.Headers.Add("In-Reply-To", $InReplyTo) }
            if ($References) { $mail.Headers.Add("References", $References) }

            if ($InlineImagePath -and $InlineCid -and (Test-Path $InlineImagePath)) {
                $htmlView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($HtmlBody, $null, "text/html")
                $linkedRes = New-Object System.Net.Mail.LinkedResource($InlineImagePath, "image/jpeg")
                $linkedRes.ContentId = $InlineCid
                $htmlView.LinkedResources.Add($linkedRes)
                $mail.AlternateViews.Add($htmlView)
            } else {
                $mail.Body = $HtmlBody
            }

            foreach ($aPath in $AttachmentPaths) {
                if ($aPath -and (Test-Path $aPath)) {
                    $att = New-Object System.Net.Mail.Attachment($aPath)
                    $mail.Attachments.Add($att)
                }
            }

            $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
            $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
            $smtp.EnableSsl = $true
            $smtp.Credentials = New-Object System.Net.NetworkCredential($From, $FromPassword)
            $smtp.Send($mail)

            $mail.Dispose()
            $smtp.Dispose()

            return @{ Success = $true; MessageId = $messageId; Error = $null }
        } catch {
            return @{ Success = $false; MessageId = $null; Error = $_.Exception.Message }
        }
    }

    Write-Log "Creating RunspacePool with $Threads threads..."
    $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $Threads)
    $RunspacePool.Open()

    # Dispatch a chunk of work items and collect results
    function Send-Chunk {
        param([array]$WorkItems)

        $jobs = [System.Collections.ArrayList]@()
        foreach ($wi in $WorkItems) {
            $ps = [PowerShell]::Create().AddScript($SendScriptBlock)
            [void]$ps.AddParameter("SmtpServer", $SmtpServer)
            [void]$ps.AddParameter("SmtpPort", $SmtpPort)
            [void]$ps.AddParameter("DomainFQDN", $DomainFQDN)
            [void]$ps.AddParameter("From", $wi.From)
            [void]$ps.AddParameter("FromPassword", $wi.FromPassword)
            [void]$ps.AddParameter("To", $wi.To)
            [void]$ps.AddParameter("Subject", $wi.Subject)
            [void]$ps.AddParameter("HtmlBody", $wi.HtmlBody)
            [void]$ps.AddParameter("AttachmentPaths", $wi.AttachmentPaths)
            [void]$ps.AddParameter("InlineImagePath", $wi.InlineImagePath)
            [void]$ps.AddParameter("InlineCid", $wi.InlineCid)
            [void]$ps.AddParameter("InReplyTo", $wi.InReplyTo)
            [void]$ps.AddParameter("References", $wi.References)
            $ps.RunspacePool = $RunspacePool
            $handle = $ps.BeginInvoke()
            [void]$jobs.Add(@{ PS = $ps; Handle = $handle; WorkItem = $wi })
        }

        # Collect results
        $results = [System.Collections.ArrayList]@()
        foreach ($job in $jobs) {
            try {
                $res = $job.PS.EndInvoke($job.Handle)
                if ($res -and $res.Count -gt 0) {
                    [void]$results.Add(@{
                        Success   = $res[0].Success
                        MessageId = $res[0].MessageId
                        Error     = $res[0].Error
                        WorkItem  = $job.WorkItem
                    })
                } else {
                    [void]$results.Add(@{ Success = $false; MessageId = $null; Error = "No result"; WorkItem = $job.WorkItem })
                }
            } catch {
                [void]$results.Add(@{ Success = $false; MessageId = $null; Error = $_.Exception.Message; WorkItem = $job.WorkItem })
            } finally {
                $job.PS.Dispose()
            }
        }

        return $results
    }

    Write-Log "Starting parallel email send ($Threads threads, chunk size $ChunkSize)..."
    Write-Log ""

    #=================================================================
    # PHASE 4a: SEND NEW MESSAGES
    #=================================================================
    Write-Log "--- Phase 4a: New Messages (target: $NewMessageTarget) ---"

    while ($newSent -lt $NewMessageTarget -and $estimatedSize -lt $TargetSizeGB) {
        # Build a chunk of work items
        $chunk = [System.Collections.ArrayList]@()
        $remaining = [math]::Min($ChunkSize, $NewMessageTarget - $newSent)

        for ($ci = 0; $ci -lt $remaining; $ci++) {
            $userIdx = ($newSent + $ci) % $UserCount
            $sender = $Users[$userIdx]
            $recipientCount = Get-Random -Minimum 1 -Maximum 4
            $recipients = Get-RandomRecipients -SenderUPN $sender.UPN -Count $recipientCount

            $inlineCid = $null; $inlineImagePath = $null
            if ((Get-Random -Minimum 1 -Maximum 101) -le $InlineImagePercent -and $JpgFiles.Count -gt 0) {
                $inlineImagePath = ($JpgFiles | Get-Random).FullName
                $inlineCid = "inlineimg_$(Get-Random)"
            }

            $htmlBody = Get-RandomHtmlBody -SenderName $sender.DisplayName -SenderEmail $sender.UPN `
                                            -RecipientName $recipients[0].DisplayName -InlineImageCid $inlineCid
            $attachPaths = Get-RandomAttachmentPaths
            $subject = $EmailSubjects | Get-Random

            [void]$chunk.Add(@{
                From            = $sender.UPN
                FromPassword    = $sender.Password
                To              = @($recipients | ForEach-Object { $_.UPN })
                Subject         = $subject
                HtmlBody        = $htmlBody
                AttachmentPaths = $attachPaths
                InlineImagePath = $inlineImagePath
                InlineCid       = $inlineCid
                InReplyTo       = $null
                References      = $null
                # Metadata for tracking
                SenderName      = $sender.DisplayName
                RecipientUPN    = $recipients[0].UPN
                RecipientName   = $recipients[0].DisplayName
            })
        }

        # Dispatch chunk in parallel
        $results = Send-Chunk -WorkItems $chunk.ToArray()

        # Process results
        $chunkOk = 0; $chunkFail = 0
        foreach ($r in $results) {
            if ($r.Success) {
                $chunkOk++
                $newSent++
                $emailsSent++

                # Store for threading
                [void]$ThreadData.messages.Add(@{
                    MessageId     = $r.MessageId
                    Subject       = $r.WorkItem.Subject
                    SenderUPN     = $r.WorkItem.From
                    SenderName    = $r.WorkItem.SenderName
                    RecipientUPN  = $r.WorkItem.RecipientUPN
                    RecipientName = $r.WorkItem.RecipientName
                })

                # Track size
                $emailSize = Get-EstimatedEmailSize -Body $r.WorkItem.HtmlBody -AttachPaths ($r.WorkItem.AttachmentPaths + @($r.WorkItem.InlineImagePath) | Where-Object { $_ })
                $estimatedSize += ($emailSize / 1073741824)
            } else {
                $chunkFail++
                $newSent++  # skip forward
            }
        }

        # Progress
        $elapsed = (Get-Date) - $startTime
        $rate = if ($elapsed.TotalMinutes -gt 0) { [math]::Round($emailsSent / $elapsed.TotalMinutes) } else { 0 }
        Write-Log "  Chunk done: $chunkOk ok, $chunkFail failed | Total: $emailsSent sent | ~$([math]::Round($estimatedSize, 2)) GB | $rate/min"

        if ($chunkFail -gt 0) {
            $sampleErr = ($results | Where-Object { -not $_.Success } | Select-Object -First 1).Error
            Write-Log "  Sample error: $sampleErr" "WARN"
        }

        # Save state
        $State.emailsSent = $emailsSent
        $State.newMessagesSent = $newSent
        $State.estimatedSizeGB = [math]::Round($estimatedSize, 4)
        Save-State $State
        if ($emailsSent % 2000 -lt $ChunkSize) { Save-Threads $ThreadData }
    }

    Write-Log "Phase 4a complete: $newSent new messages sent"
    Save-Threads $ThreadData

    #=================================================================
    # PHASE 4b: SEND REPLIES
    #=================================================================
    Write-Log "--- Phase 4b: Replies (target: $ReplyTarget) ---"

    while ($replySent -lt $ReplyTarget -and $estimatedSize -lt $TargetSizeGB -and $ThreadData.messages.Count -gt 0) {
        $chunk = [System.Collections.ArrayList]@()
        $remaining = [math]::Min($ChunkSize, $ReplyTarget - $replySent)

        for ($ci = 0; $ci -lt $remaining; $ci++) {
            $origMsg = $ThreadData.messages | Get-Random
            $replySenderUPN = $origMsg.RecipientUPN
            $replySenderCred = $UserCredMap[$replySenderUPN]
            if (-not $replySenderCred) { continue }

            $replyText = $TextSnippets | Get-Random
            $sig = Get-HtmlSignature -DisplayName $origMsg.RecipientName -Email $replySenderUPN

            $replyHtml = @"
<html><head><meta charset="utf-8"></head>
<body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#333;line-height:1.6;">
<p>$replyText</p>
$sig
<hr style="border:none;border-top:1px solid #ccc;margin:15px 0;"/>
<p style="color:#666;font-size:12px;">On $(Get-Date -Format 'ddd, MMM d yyyy HH:mm'), $($origMsg.SenderName) &lt;$($origMsg.SenderUPN)&gt; wrote:</p>
<blockquote style="border-left:2px solid #4472C4;margin:10px 0;padding-left:10px;color:#555;">
$($TextSnippets | Get-Random)
</blockquote>
</body></html>
"@
            $attachPaths = Get-RandomAttachmentPaths

            [void]$chunk.Add(@{
                From            = $replySenderUPN
                FromPassword    = $replySenderCred.Password
                To              = @($origMsg.SenderUPN)
                Subject         = "Re: $($origMsg.Subject)"
                HtmlBody        = $replyHtml
                AttachmentPaths = $attachPaths
                InlineImagePath = $null
                InlineCid       = $null
                InReplyTo       = $origMsg.MessageId
                References      = $origMsg.MessageId
                SenderName      = $origMsg.RecipientName
                RecipientUPN    = $origMsg.SenderUPN
                RecipientName   = $origMsg.SenderName
            })
        }

        if ($chunk.Count -eq 0) { break }
        $results = Send-Chunk -WorkItems $chunk.ToArray()

        $chunkOk = 0; $chunkFail = 0
        foreach ($r in $results) {
            if ($r.Success) {
                $chunkOk++
                $replySent++
                $emailsSent++
                [void]$ThreadData.messages.Add(@{
                    MessageId = $r.MessageId; Subject = $r.WorkItem.Subject
                    SenderUPN = $r.WorkItem.From; SenderName = $r.WorkItem.SenderName
                    RecipientUPN = $r.WorkItem.RecipientUPN; RecipientName = $r.WorkItem.RecipientName
                })
                $emailSize = Get-EstimatedEmailSize -Body $r.WorkItem.HtmlBody -AttachPaths $r.WorkItem.AttachmentPaths
                $estimatedSize += ($emailSize / 1073741824)
            } else { $chunkFail++; $replySent++ }
        }

        $elapsed = (Get-Date) - $startTime
        $rate = if ($elapsed.TotalMinutes -gt 0) { [math]::Round($emailsSent / $elapsed.TotalMinutes) } else { 0 }
        Write-Log "  Chunk: $chunkOk ok, $chunkFail fail | Total: $emailsSent | ~$([math]::Round($estimatedSize, 2)) GB | $rate/min"

        $State.emailsSent = $emailsSent; $State.repliesSent = $replySent
        $State.estimatedSizeGB = [math]::Round($estimatedSize, 4)
        Save-State $State
    }

    Write-Log "Phase 4b complete: $replySent replies sent"

    #=================================================================
    # PHASE 4c: SEND FORWARDS
    #=================================================================
    Write-Log "--- Phase 4c: Forwards (target: $ForwardTarget) ---"

    while ($fwdSent -lt $ForwardTarget -and $estimatedSize -lt $TargetSizeGB -and $ThreadData.messages.Count -gt 0) {
        $chunk = [System.Collections.ArrayList]@()
        $remaining = [math]::Min($ChunkSize, $ForwardTarget - $fwdSent)

        for ($ci = 0; $ci -lt $remaining; $ci++) {
            $origMsg = $ThreadData.messages | Get-Random
            $fwdSenderUPN = $origMsg.RecipientUPN
            $fwdSenderCred = $UserCredMap[$fwdSenderUPN]
            if (-not $fwdSenderCred) { continue }

            $fwdRecipient = Get-RandomRecipients -SenderUPN $fwdSenderUPN -Count 1
            $fwdText = $TextSnippets | Get-Random
            $sig = Get-HtmlSignature -DisplayName $origMsg.RecipientName -Email $fwdSenderUPN

            $fwdHtml = @"
<html><head><meta charset="utf-8"></head>
<body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#333;line-height:1.6;">
<p style="color:#4472C4;font-weight:bold;">FYI — see below:</p>
<p>$fwdText</p>
$sig
<hr style="border:none;border-top:1px solid #ccc;margin:15px 0;"/>
<p style="color:#666;font-size:12px;"><b>From:</b> $($origMsg.SenderName) &lt;$($origMsg.SenderUPN)&gt;<br/>
<b>Subject:</b> $($origMsg.Subject)</p>
<div style="padding-left:10px;border-left:2px solid #ddd;">
$($TextSnippets | Get-Random)
</div>
</body></html>
"@
            $attachPaths = Get-RandomAttachmentPaths

            [void]$chunk.Add(@{
                From            = $fwdSenderUPN
                FromPassword    = $fwdSenderCred.Password
                To              = @($fwdRecipient[0].UPN)
                Subject         = "FW: $($origMsg.Subject)"
                HtmlBody        = $fwdHtml
                AttachmentPaths = $attachPaths
                InlineImagePath = $null
                InlineCid       = $null
                InReplyTo       = $null
                References      = $null
                SenderName      = $origMsg.RecipientName
                RecipientUPN    = $fwdRecipient[0].UPN
                RecipientName   = $fwdRecipient[0].DisplayName
            })
        }

        if ($chunk.Count -eq 0) { break }
        $results = Send-Chunk -WorkItems $chunk.ToArray()

        $chunkOk = 0; $chunkFail = 0
        foreach ($r in $results) {
            if ($r.Success) {
                $chunkOk++; $fwdSent++; $emailsSent++
                $emailSize = Get-EstimatedEmailSize -Body $r.WorkItem.HtmlBody -AttachPaths $r.WorkItem.AttachmentPaths
                $estimatedSize += ($emailSize / 1073741824)
            } else { $chunkFail++; $fwdSent++ }
        }

        $elapsed = (Get-Date) - $startTime
        $rate = if ($elapsed.TotalMinutes -gt 0) { [math]::Round($emailsSent / $elapsed.TotalMinutes) } else { 0 }
        Write-Log "  Chunk: $chunkOk ok, $chunkFail fail | Total: $emailsSent | ~$([math]::Round($estimatedSize, 2)) GB | $rate/min"

        $State.emailsSent = $emailsSent; $State.forwardsSent = $fwdSent
        $State.estimatedSizeGB = [math]::Round($estimatedSize, 4)
        Save-State $State
    }

    Write-Log "Phase 4c complete: $fwdSent forwards sent"

    #=================================================================
    # Extra round if target not reached
    #=================================================================
    if ($estimatedSize -lt $TargetSizeGB) {
        Write-Log "--- Extra round: $([math]::Round($estimatedSize, 2))GB / ${TargetSizeGB}GB — sending more with large attachments ---"

        while ($estimatedSize -lt $TargetSizeGB) {
            $chunk = [System.Collections.ArrayList]@()
            $largeFiles = @($AllAttachments | Sort-Object Length -Descending | Select-Object -First 10 | ForEach-Object { $_.FullName })

            for ($ci = 0; $ci -lt $ChunkSize; $ci++) {
                $userIdx = ($emailsSent + $ci) % $UserCount
                $sender = $Users[$userIdx]
                $recipients = Get-RandomRecipients -SenderUPN $sender.UPN -Count (Get-Random -Minimum 1 -Maximum 4)

                $inlineCid = $null; $inlineImagePath = $null
                if ((Get-Random -Minimum 1 -Maximum 101) -le $InlineImagePercent -and $JpgFiles.Count -gt 0) {
                    $inlineImagePath = ($JpgFiles | Get-Random).FullName
                    $inlineCid = "inlineimg_$(Get-Random)"
                }

                $htmlBody = Get-RandomHtmlBody -SenderName $sender.DisplayName -SenderEmail $sender.UPN `
                                                -RecipientName $recipients[0].DisplayName -InlineImageCid $inlineCid

                # Bias towards larger attachments
                $attachPaths = if (($TargetSizeGB - $estimatedSize) -gt 5) {
                    @($largeFiles | Get-Random -Count (Get-Random -Minimum 1 -Maximum 4))
                } else {
                    Get-RandomAttachmentPaths
                }

                [void]$chunk.Add(@{
                    From = $sender.UPN; FromPassword = $sender.Password
                    To = @($recipients | ForEach-Object { $_.UPN })
                    Subject = ($EmailSubjects | Get-Random); HtmlBody = $htmlBody
                    AttachmentPaths = $attachPaths; InlineImagePath = $inlineImagePath
                    InlineCid = $inlineCid; InReplyTo = $null; References = $null
                    SenderName = $sender.DisplayName
                    RecipientUPN = $recipients[0].UPN; RecipientName = $recipients[0].DisplayName
                })
            }

            $results = Send-Chunk -WorkItems $chunk.ToArray()

            $chunkOk = 0
            foreach ($r in $results) {
                if ($r.Success) {
                    $chunkOk++; $newSent++; $emailsSent++
                    $emailSize = Get-EstimatedEmailSize -Body $r.WorkItem.HtmlBody -AttachPaths ($r.WorkItem.AttachmentPaths + @($r.WorkItem.InlineImagePath) | Where-Object { $_ })
                    $estimatedSize += ($emailSize / 1073741824)
                } else { $emailsSent++ }
            }

            $elapsed = (Get-Date) - $startTime
            $rate = if ($elapsed.TotalMinutes -gt 0) { [math]::Round($emailsSent / $elapsed.TotalMinutes) } else { 0 }
            Write-Log "  Extra: $chunkOk ok | Total: $emailsSent | ~$([math]::Round($estimatedSize, 2)) GB | $rate/min"

            $State.emailsSent = $emailsSent; $State.newMessagesSent = $newSent
            $State.estimatedSizeGB = [math]::Round($estimatedSize, 4)
            Save-State $State
        }
    }

    # Cleanup RunspacePool
    $RunspacePool.Close()
    $RunspacePool.Dispose()
    Write-Log "RunspacePool closed"

    # Final state save
    Save-Threads $ThreadData
    $State.emailsSent = $emailsSent
    $State.newMessagesSent = $newSent
    $State.repliesSent = $replySent
    $State.forwardsSent = $fwdSent
    $State.estimatedSizeGB = [math]::Round($estimatedSize, 4)
    $State.phase = 5
    Save-State $State

    Write-Host ""
    Write-Host "--- Phase 4 Complete ---" -ForegroundColor Green
    Write-Host ""
}

#####################################################################
# PHASE 5: REPORT
#####################################################################

if ($StartPhase -le 5) {
    Write-Host "=== PHASE 5: Final Report ===" -ForegroundColor Yellow
    Write-Host ""

    # Ensure DC is set (in case we're resuming from Phase 5)
    if (-not $script:TargetDC) {
        $script:TargetDC = (Get-ADDomain).PDCEmulator
    }
    $DC = $script:TargetDC

    $State = Get-State
    $endTime = Get-Date
    $totalElapsed = if ($State.startTime) { $endTime - [datetime]$State.startTime } else { [timespan]::Zero }

    # Try to get actual database size
    $actualDbSize = "N/A"
    try {
        if (Get-Command Get-MailboxDatabase -ErrorAction SilentlyContinue) {
            $dbInfo = Get-MailboxDatabase -Status -DomainController $DC | Select-Object Name, DatabaseSize
            $actualDbSize = ($dbInfo | ForEach-Object { "$($_.Name): $($_.DatabaseSize)" }) -join "; "
        }
    } catch { }

    # Try to get sample mailbox stats
    $sampleStats = "N/A"
    try {
        if (Get-Command Get-MailboxStatistics -ErrorAction SilentlyContinue) {
            $DomainFQDN = (Get-ADDomain -Server $DC).DNSRoot
            $stats = Get-MailboxStatistics "${UserPrefix}001@$DomainFQDN" -DomainController $DC -ErrorAction SilentlyContinue
            if ($stats) {
                $sampleStats = "ItemCount: $($stats.ItemCount) | TotalSize: $($stats.TotalItemSize)"
            }
        }
    } catch { }

    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  GENERATION COMPLETE" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Users Created:       $($State.usersCreated)" -ForegroundColor White
    Write-Host "  Total Emails Sent:   $($State.emailsSent)" -ForegroundColor White
    Write-Host "    New Messages:      $($State.newMessagesSent)" -ForegroundColor White
    Write-Host "    Replies:           $($State.repliesSent)" -ForegroundColor White
    Write-Host "    Forwards:          $($State.forwardsSent)" -ForegroundColor White
    Write-Host "  Estimated Size:      $([math]::Round($State.estimatedSizeGB, 2)) GB" -ForegroundColor White
    Write-Host "  Actual DB Size:      $actualDbSize" -ForegroundColor White
    Write-Host "  Sample Mailbox:      $sampleStats" -ForegroundColor White
    Write-Host "  Elapsed Time:        $($totalElapsed.ToString('d\.hh\:mm\:ss'))" -ForegroundColor White
    Write-Host ""
    Write-Host "  Credentials File:    $CredsFile" -ForegroundColor Gray
    Write-Host "  Log File:            $LogFile" -ForegroundColor Gray
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""

    # Export report CSV
    $reportData = [PSCustomObject]@{
        Timestamp = $endTime.ToString("yyyy-MM-dd HH:mm:ss")
        UsersCreated = $State.usersCreated
        TotalEmailsSent = $State.emailsSent
        NewMessages = $State.newMessagesSent
        Replies = $State.repliesSent
        Forwards = $State.forwardsSent
        EstimatedSizeGB = [math]::Round($State.estimatedSizeGB, 2)
        ActualDBSize = $actualDbSize
        ElapsedTime = $totalElapsed.ToString('d\.hh\:mm\:ss')
        TargetSizeGB = $TargetSizeGB
        UserCount = $UserCount
    }
    $reportData | Export-Csv -Path $ReportFile -NoTypeInformation -Encoding UTF8
    Write-Log "Report saved to: $ReportFile" "OK"

    $State.phase = 5
    $State.completed = $true
    Save-State $State
}

Write-Host "Done." -ForegroundColor Green
