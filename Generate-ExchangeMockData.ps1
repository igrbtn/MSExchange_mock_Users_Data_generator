#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange Mock Data Generator — creates 300 users and ~100GB of realistic email data.

.DESCRIPTION
    Runs on Exchange Server 2019 with Domain Admin + Organization Management.
    Phase 1: Create 300 mailbox-enabled users
    Phase 2: Validate SMTP relay on localhost:25
    Phase 3: Download/generate sample attachments (JPG, TXT, RTF)
    Phase 4: Provision default mailbox folders via EWS login, then send ~34,000 emails via authenticated SMTP (new, reply, forward) with HTML formatting and inline images
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
    [int]$UserCount = 0,
    [string]$UserPrefix = "mockuser",
    [string]$MockUsersOU = "",
    [string]$UserPassword = "",
    [int]$PasswordLength = 12,
    [int]$Threads = 10,
    [int]$ChunkSize = 100,
    [int]$MaxAttachmentSizeMB = 10,
    [int]$InlineImagePercent = 30,
    [string]$Database = "",
    [string]$SmtpServer = "",
    [string]$EwsServer = "",
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
        database = ""
        smtpServer = ""
        ewsServer = ""
        impersonationReady = $false
        attachmentsReady = $false
        foldersProvisioned = $false
        contactsCreated = $false
        calendarCreated = $false
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
    # English/American
    "James","John","Robert","David","Michael","William","Thomas","Daniel","Richard","Joseph",
    "Andrew","Alex","Ryan","Nathan","Brian","Kevin","Eric","Scott","Patrick","Timothy",
    "Mary","Jennifer","Linda","Susan","Jessica","Sarah","Karen","Lisa","Nancy","Betty",
    # Russian
    "Dmitry","Sergei","Alexei","Nikolai","Vladimir","Andrei","Mikhail","Pavel","Igor","Oleg",
    "Olga","Elena","Natasha","Svetlana","Irina","Tatiana","Anastasia","Yulia","Marina","Vera",
    # Spanish/Portuguese
    "Carlos","Miguel","Diego","Alejandro","Fernando","Rafael","Sergio","Pablo","Emilio","Mateo",
    "Maria","Sofia","Isabella","Camila","Valentina","Gabriela","Paula","Laura","Carmen","Andrea",
    # French
    "Pierre","Jean","Louis","Antoine","Nicolas","Philippe","Etienne","Francois","Henri","Remy",
    "Colette","Juliette","Amelie","Margaux","Celine","Claire","Eloise","Isabelle","Simone","Adele",
    # German
    "Hans","Klaus","Franz","Wolfgang","Dieter","Markus","Stefan","Tobias","Lukas","Felix",
    "Helga","Greta","Ingrid","Heidi","Ursula","Anke","Sabine","Monika","Petra","Martina",
    # Chinese
    "Wei","Chen","Jun","Liang","Hao","Jian","Ming","Tao","Xin","Feng",
    "Mei","Ling","Xia","Yan","Hui","Na","Jing","Fang","Yun","Li",
    # Japanese
    "Hiroshi","Yuki","Takeshi","Kenji","Ryu","Kaito","Shota","Daichi","Haruto","Ren",
    "Sakura","Keiko","Hana","Rin","Yui","Aoi","Mio","Emi","Nana","Akari",
    # Arabic
    "Ahmed","Mohammed","Ali","Omar","Hassan","Ibrahim","Yusuf","Khalid","Tariq","Faisal",
    "Fatima","Aisha","Layla","Noor","Zara","Amira","Hala","Dina","Sara","Leila",
    # Indian
    "Priya","Deepa","Ananya","Kavya","Meera","Riya","Nisha","Sita","Pooja","Aditi",
    "Arjun","Vikram","Ravi","Suresh","Anil","Rahul","Amit","Sanjay","Kiran","Dev",
    # Korean
    "Minho","Jisoo","Hyun","Seojin","Taeyang","Dohyun","Jiho","Yunho","Sungho","Woojin",
    "Minji","Yuna","Soyeon","Haeun","Chaewon","Jiwoo","Dahye","Eunbi","Nayoung","Subin",
    # Italian
    "Marco","Luca","Giuseppe","Alessandro","Matteo","Lorenzo","Davide","Riccardo","Giorgio","Fabio",
    "Giulia","Francesca","Chiara","Alessia","Valentina","Elisa","Silvia","Roberta","Monica","Paola",
    # Nordic/Scandinavian
    "Erik","Lars","Magnus","Sven","Olaf","Bjorn","Nils","Torsten","Leif","Ragnar",
    "Astrid","Ingrid","Sigrid","Frida","Elsa","Kristin","Birgit","Solveig","Liv","Maja"
)

$LastNames = @(
    # English
    "Smith","Johnson","Williams","Brown","Jones","Davis","Miller","Wilson","Moore","Taylor",
    "Anderson","White","Harris","Clark","Lewis","Walker","Hall","Young","Allen","King",
    # Russian
    "Ivanov","Petrov","Sokolov","Kuznetsov","Popov","Volkov","Morozov","Novikov","Kozlov","Lebedev",
    "Sorokin","Pavlov","Semenov","Egorov","Fedorov","Orlov","Belov","Zakharov","Voronov","Gusev",
    # Spanish/Portuguese
    "Garcia","Martinez","Lopez","Hernandez","Gonzalez","Rodriguez","Perez","Sanchez","Ramirez","Torres",
    "Silva","Santos","Oliveira","Pereira","Costa","Ferreira","Almeida","Souza","Lima","Ribeiro",
    # French
    "Martin","Bernard","Dubois","Moreau","Laurent","Simon","Michel","Leroy","Roux","Fontaine",
    # German
    "Mueller","Schmidt","Schneider","Fischer","Weber","Meyer","Wagner","Becker","Schulz","Hoffmann",
    "Bauer","Koch","Richter","Klein","Kraus","Werner","Lehmann","Braun","Zimmermann","Hartmann",
    # Chinese
    "Wang","Li","Zhang","Liu","Chen","Yang","Huang","Wu","Zhou","Xu",
    "Sun","Ma","Zhu","Gao","Lin","Zhao","Deng","Feng","Luo","Tang",
    # Japanese
    "Tanaka","Yamamoto","Suzuki","Takahashi","Watanabe","Ito","Nakamura","Kobayashi","Kato","Yoshida",
    "Yamada","Sasaki","Yamaguchi","Matsumoto","Inoue","Kimura","Shimizu","Hayashi","Mori","Saito",
    # Arabic
    "Al-Said","Hassan","Ibrahim","Ali","Mahmoud","Yusuf","Mustafa","Rashid","Saleh","Hamid",
    "Osman","Abbas","Nasser","Farid","Kareem","Saeed","Khalil","Bakr","Mansour","Darwish",
    # Indian
    "Sharma","Patel","Gupta","Singh","Kumar","Reddy","Joshi","Verma","Nair","Rao",
    "Shah","Pillai","Iyer","Mishra","Tiwari","Bhat","Desai","Mehta","Chopra","Kapoor",
    # Korean
    "Kim","Lee","Park","Choi","Jung","Kang","Yoon","Song","Lim","Han",
    # Italian
    "Rossi","Russo","Ferrari","Esposito","Bianchi","Romano","Colombo","Ricci","Marino","Greco",
    # Nordic
    "Johansson","Lindberg","Eriksson","Nilsson","Larsson","Olsson","Andersen","Dahl","Bakken","Berg"
)

#####################################################################
# USER PROFILE DATA (randomized per user)
#####################################################################

$Departments = @(
    "IT","Finance","Human Resources","Marketing","Sales","Engineering",
    "Operations","Legal","Customer Support","Research & Development",
    "Procurement","Quality Assurance","Business Development","Administration",
    "Product Management","Data Analytics","Security","Compliance",
    "Logistics","Communications","Training","Facilities"
)

$JobTitles = @(
    "Analyst","Senior Analyst","Manager","Senior Manager","Director",
    "Vice President","Team Lead","Specialist","Coordinator","Administrator",
    "Engineer","Senior Engineer","Principal Engineer","Architect","Consultant",
    "Senior Consultant","Associate","Executive","Officer","Supervisor",
    "Developer","Project Manager","Program Manager","Account Manager","Advisor",
    "Technician","Assistant","Planner","Strategist","Controller"
)

$Offices = @(
    "HQ-101","HQ-102","HQ-201","HQ-202","HQ-301","HQ-302",
    "Building A, Room 101","Building A, Room 205","Building A, Room 310",
    "Building B, Room 102","Building B, Room 204","Building B, Room 306",
    "Building C, Room 103","Building C, Room 201","Building C, Room 305",
    "Tower 1, Floor 5","Tower 1, Floor 10","Tower 1, Floor 15",
    "Tower 2, Floor 3","Tower 2, Floor 8","Tower 2, Floor 12",
    "Remote","Remote","Remote","Remote"
)

$Cities = @(
    "New York","London","Moscow","Tokyo","Berlin","Paris","Dubai",
    "Mumbai","Shanghai","Sydney","Toronto","Singapore","Seoul",
    "Rome","Madrid","Zurich","Amsterdam","Stockholm","Vienna","Prague",
    "Istanbul","Bangkok","Sao Paulo","Mexico City","Lagos","Cairo"
)

$Countries = @(
    "US","GB","RU","JP","DE","FR","AE","IN","CN","AU",
    "CA","SG","KR","IT","ES","CH","NL","SE","AT","CZ",
    "TR","TH","BR","MX","NG","EG"
)

$Companies = @(
    "Global Solutions Ltd","TechVista Corp","DataBridge Inc","Nexus Innovations",
    "Meridian Systems","Apex Consulting","Quantum Networks","Horizon Group",
    "Pinnacle Services","CoreTech International","Vector Dynamics","Summit Partners",
    "Atlas Enterprises","Zenith Holdings","Vertex Analytics","Catalyst Engineering",
    "Streamline Corp","Fusion Technologies","Vanguard Solutions","Prism Consulting"
)

$StreetAddresses = @(
    "100 Main Street","200 Broadway","42 Park Avenue","15 Technology Drive",
    "350 Enterprise Blvd","77 Innovation Way","500 Commerce Street",
    "1200 Corporate Center","88 Business Park Road","250 Financial Drive",
    "10 Market Square","300 Industrial Parkway","60 Riverside Drive",
    "175 Liberty Avenue","425 Global Way","90 Silicon Road"
)

$PostalCodes = @(
    "10001","SW1A 1AA","101000","100-0001","10115","75001","00000",
    "400001","200000","2000","M5V 2T6","018956","06164","00100",
    "28001","8001","1012","111 21","1010","110 00"
)

$Descriptions = @(
    "Employee — Mock user for testing and migration validation",
    "Staff member — Generated for Exchange mock data project",
    "Test account — Part of bulk user generation for lab environment",
    "Mock user — Created for mailbox migration testing purposes",
    "Lab account — Used for Exchange data generation testing"
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
    # English — business
    "Quarterly Report Q4 2025", "Meeting Request: Project Update",
    "Action Required: Review Document", "Weekly Status Update",
    "Team Sync — Priorities", "Follow-up: Client Feedback",
    "Updated Proposal Draft", "Infrastructure Maintenance Notice",
    "New Policy Guidelines", "Training Schedule Update",
    "Vendor Evaluation Results", "Project Timeline Revision",
    "Security Audit Findings", "Product Launch Roadmap",
    "Customer Survey Results", "Office Relocation Plan",
    "Annual Review Preparation", "Partnership Opportunity",
    "Technical Specification Review", "Compliance Update",
    "Budget Forecast FY2026", "Team Building Event",
    "Service Level Agreement Draft", "Data Migration Status",
    "Performance Metrics Report", "Risk Assessment Summary",
    "Recruitment Update", "IT Support Ticket Summary",
    "Network Upgrade Schedule", "Firewall Rule Change Request",
    "Access Permissions Review", "Deployment Checklist — Production",
    "Backup Verification Report", "Inventory Reconciliation Q1",
    "Travel Authorization Request", "Expense Report — January",
    "Contract Renewal Discussion", "Software License Audit",
    "Incident Report #4521", "Change Management Board Agenda",
    "Disaster Recovery Test Results", "Server Decommission Plan",
    "Onboarding Checklist — New Hire", "Exit Interview Summary",
    "Capacity Planning — Q2 2026", "SSL Certificate Expiration Notice",
    "Patch Tuesday Update", "VPN Configuration Change",
    "Print Queue Migration", "Phone System Upgrade",
    "Holiday Schedule 2026", "Parking Lot Assignment Update",
    "Health & Safety Inspection Report", "Fire Drill Schedule",
    "Key Card Access Request", "Visitor Badge Policy Update",
    "Equipment Return Form", "Asset Tag Verification",
    "Monthly KPI Dashboard", "Customer Escalation — Priority",
    "Account Reconciliation Notice", "Wire Transfer Confirmation",
    "Invoice Discrepancy — PO#38291", "Supplier Payment Schedule",
    # Russian
    "Отчёт за квартал", "Приглашение на совещание", "Обновление проекта",
    "Заявка на отпуск", "Согласование бюджета", "План миграции серверов",
    "Результаты аудита безопасности", "Запрос на доступ к системе",
    "Обновление политики ИБ", "Акт выполненных работ",
    # French
    "Rapport trimestriel", "Invitation reunion", "Mise a jour du projet",
    "Demande de conge", "Approbation du budget", "Plan de migration",
    "Rapport d'audit de securite", "Demande d'acces systeme",
    # Spanish
    "Informe trimestral", "Solicitud de reunion", "Actualizacion del proyecto",
    "Solicitud de vacaciones", "Aprobacion de presupuesto",
    "Plan de migracion de servidores", "Informe de auditoria",
    # German
    "Quartalsbericht", "Besprechungseinladung", "Projekt-Update",
    "Urlaubsantrag", "Budgetfreigabe", "Server-Migrationsplan",
    "Sicherheitsaudit Ergebnisse", "Systemzugangsanfrage",
    # Japanese
    "四半期報告書", "会議の招待", "プロジェクト更新",
    "休暇申請", "予算承認依頼", "サーバー移行計画",
    # Chinese
    "季度报告", "会议邀请", "项目更新进展",
    "休假申请", "预算审批", "服务器迁移计划",
    # Portuguese
    "Relatorio trimestral", "Convite para reuniao", "Atualizacao do projeto",
    "Solicitacao de ferias", "Aprovacao de orcamento", "Plano de migracao",
    # Italian
    "Rapporto trimestrale", "Invito alla riunione", "Aggiornamento progetto",
    "Richiesta ferie", "Approvazione budget", "Piano di migrazione server",
    # Korean
    "분기 보고서", "회의 초대", "프로젝트 업데이트", "휴가 신청", "예산 승인 요청",
    # Arabic
    "تقرير ربع سنوي", "دعوة اجتماع", "تحديث المشروع", "طلب إجازة", "الموافقة على الميزانية"
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
    # ---- Interactive prompts (if not provided via parameters) ----

    if ($UserCount -eq 0) {
        Write-Host ""
        $input = Read-Host "  How many users to create? [300]"
        $UserCount = if ([string]::IsNullOrWhiteSpace($input)) { 300 } else { [int]$input }
    }

    if ([string]::IsNullOrEmpty($MockUsersOU)) {
        $input = Read-Host "  OU name for mock users? [MockUsers]"
        $MockUsersOU = if ([string]::IsNullOrWhiteSpace($input)) { "MockUsers" } else { $input.Trim() }
    }

    if ([string]::IsNullOrEmpty($UserPassword)) {
        Write-Host ""
        Write-Host "  Password mode:" -ForegroundColor Cyan
        Write-Host "    [1] Generate unique random password per user (default)" -ForegroundColor White
        Write-Host "    [2] Use same password for all users" -ForegroundColor White
        Write-Host ""
        $pwChoice = Read-Host "  Select (1 or 2) [1]"
        if ($pwChoice -eq '2') {
            $UserPassword = Read-Host "  Enter password for all users"
            if ([string]::IsNullOrWhiteSpace($UserPassword)) {
                Write-Host "  Empty password, falling back to random generation" -ForegroundColor Yellow
                $UserPassword = ""
            }
        }
    }

    Write-Host ""
    Write-Host "=== PHASE 1: Creating $UserCount Mock Users ===" -ForegroundColor Yellow
    Write-Host "    OU: $MockUsersOU" -ForegroundColor Gray
    Write-Host "    Password: $(if ($UserPassword) { 'Same for all' } else { 'Random per user' })" -ForegroundColor Gray
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

    # ---- Select mailbox database ----
    # Use saved database from state on resume, or -Database param, or ask interactively
    if ($State.database) { $Database = $State.database }
    if ([string]::IsNullOrEmpty($Database)) {
        $allDBs = @(Get-MailboxDatabase -DomainController $DC -Status | Select-Object Name, ServerName, DatabaseSize | Sort-Object Name)
        if ($allDBs.Count -eq 0) {
            Write-Log "No mailbox databases found" "ERROR"
            exit 1
        } elseif ($allDBs.Count -eq 1) {
            $MDB = $allDBs[0].Name
            Write-Log "Single database detected: $MDB" "OK"
        } else {
            Write-Host ""
            Write-Host "  Available Mailbox Databases:" -ForegroundColor Cyan
            for ($di = 0; $di -lt $allDBs.Count; $di++) {
                $dbSize = if ($allDBs[$di].DatabaseSize) { $allDBs[$di].DatabaseSize.ToString() } else { "N/A" }
                Write-Host "    [$($di + 1)] $($allDBs[$di].Name)  (Server: $($allDBs[$di].ServerName), Size: $dbSize)" -ForegroundColor White
            }
            Write-Host ""
            $dbChoice = Read-Host "  Select database number (1-$($allDBs.Count))"
            $dbIdx = [int]$dbChoice - 1
            if ($dbIdx -lt 0 -or $dbIdx -ge $allDBs.Count) {
                Write-Log "Invalid choice, using first database" "WARN"
                $dbIdx = 0
            }
            $MDB = $allDBs[$dbIdx].Name
            Write-Log "Selected database: $MDB" "OK"
        }
    } else {
        $MDB = $Database
        Write-Log "Using specified database: $MDB"
    }
    $State.database = $MDB
    Save-State $State

    # Prepare credentials CSV
    $CredsData = @()
    $Created = 0
    $Skipped = 0

    # Pre-shuffle names for unique combinations (seed with user count for reproducibility)
    $shuffledFirst = $FirstNames | Sort-Object { Get-Random }
    $shuffledLast = $LastNames | Sort-Object { Get-Random }

    for ($i = 1; $i -le $UserCount; $i++) {
        $num = $i.ToString("D3")
        $alias = "$UserPrefix$num"
        $upn = "$alias@$DomainFQDN"

        # Pick name — use shuffled arrays for uniqueness, wrap around if > array size
        $firstName = $shuffledFirst[(($i - 1) % $shuffledFirst.Count)]
        $lastName = $shuffledLast[(($i - 1) % $shuffledLast.Count)]
        $displayName = "MockOrg $lastName $firstName"

        # Pick random profile attributes
        $dept = $Departments | Get-Random
        $title = $JobTitles | Get-Random
        $office = $Offices | Get-Random
        $cityIdx = Get-Random -Maximum $Cities.Count
        $city = $Cities[$cityIdx]
        $country = $Countries[($cityIdx % $Countries.Count)]
        $company = $Companies | Get-Random
        $streetAddr = $StreetAddresses | Get-Random
        $postalCode = $PostalCodes | Get-Random
        $description = $Descriptions | Get-Random
        $phoneExt = (Get-Random -Minimum 1000 -Maximum 9999).ToString()
        $phone = "+1-555-$phoneExt"
        $mobilePrefix = @("+1-555-","+7-916-","+44-7700-","+49-170-","+81-90-","+86-138-") | Get-Random
        $mobile = "$mobilePrefix$(Get-Random -Minimum 1000000 -Maximum 9999999)"
        $initials = "$($firstName[0])$($lastName[0])"

        # Check if already exists (pinned to DC)
        $existing = Get-Mailbox -Identity $alias -DomainController $DC -ErrorAction SilentlyContinue
        if ($existing) {
            $existingDB = $existing.Database.Name
            if ($existingDB -ne $MDB) {
                # User is on a different database — remove and recreate on the selected one
                Write-Log "  [$i/$UserCount] $alias is on '$existingDB', removing to recreate on '$MDB'..." "INFO"
                try {
                    Remove-Mailbox -Identity $alias -DomainController $DC -Permanent $true -Confirm:$false -ErrorAction Stop
                    Start-Sleep -Seconds 2
                    $existing = $null   # fall through to normal creation below
                } catch {
                    Write-Log "  [$i/$UserCount] Failed to remove ${alias}: $_" "ERROR"
                    $password = "***FAILED***"
                    $CredsData += [PSCustomObject]@{
                        Number = $i; Alias = $alias; UPN = $upn
                        DisplayName = $displayName; Password = $password; SamAccountName = $alias
                        FirstName = $firstName; LastName = $lastName
                        Department = $dept; Title = $title; Office = $office
                        City = $city; Country = $country; Company = $company
                        Phone = $phone; Mobile = $mobile
                    }
                    continue
                }
            }
        }

        if ($existing) {
            # User exists and is already on the correct database — keep it, reset password
            $Skipped++

            $password = if ($UserPassword) { $UserPassword } else { New-RandomPassword -Length $PasswordLength }
            $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
            try {
                Set-ADAccountPassword -Identity $alias -NewPassword $securePassword -Reset -Server $DC -ErrorAction Stop
                Set-ADUser -Identity $alias -ChangePasswordAtLogon $false -Server $DC -ErrorAction SilentlyContinue
            } catch {
                Write-Log "Could not reset password for ${alias}: $_" "WARN"
                $password = "***existing***"
            }

            $CredsData += [PSCustomObject]@{
                Number = $i; Alias = $alias; UPN = $upn
                DisplayName = $displayName; Password = $password; SamAccountName = $alias
                FirstName = $firstName; LastName = $lastName
                Department = $dept; Title = $title; Office = $office
                City = $city; Country = $country; Company = $company
                Phone = $phone; Mobile = $mobile
            }

            # Still update profile fields on existing users
            try {
                Set-User -Identity $alias -Department $dept -Title $title -Office $office `
                    -City $city -CountryOrRegion $country -Company $company `
                    -StreetAddress $streetAddr -PostalCode $postalCode `
                    -Phone $phone -MobilePhone $mobile -Initials $initials `
                    -Notes $description `
                    -DomainController $DC -ErrorAction SilentlyContinue
            } catch { }
            continue
        }

        $password = if ($UserPassword) { $UserPassword } else { New-RandomPassword -Length $PasswordLength }
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force

        try {
            # Name (AD CN) = DisplayName = "MockOrg LastName FirstName"
            New-Mailbox -Name $displayName `
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

            # Set additional user profile fields
            Set-User -Identity $alias `
                -Department $dept `
                -Title $title `
                -Office $office `
                -City $city `
                -CountryOrRegion $country `
                -Company $company `
                -StreetAddress $streetAddr `
                -PostalCode $postalCode `
                -Phone $phone `
                -MobilePhone $mobile `
                -Initials $initials `
                -Notes $description `
                -DomainController $DC `
                -ErrorAction SilentlyContinue

            $Created++
            Write-Log "  [$i/$UserCount] Created: $upn ($displayName) — $title, $dept" "OK"
        } catch {
            Write-Log "  [$i/$UserCount] Failed: $upn — $_" "ERROR"
            $password = "***FAILED***"
        }

        $CredsData += [PSCustomObject]@{
            Number = $i; Alias = $alias; UPN = $upn
            DisplayName = $displayName; Password = $password; SamAccountName = $alias
            FirstName = $firstName; LastName = $lastName
            Department = $dept; Title = $title; Office = $office
            City = $city; Country = $country; Company = $company
            Phone = $phone; Mobile = $mobile
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

    # ---- Detect Exchange servers ----
    Write-Log "Detecting Exchange servers..."
    $ExServers = @()
    try {
        $ExServers = @(Get-ExchangeServer | Where-Object {
            $_.ServerRole -match "Mailbox" -and $_.AdminDisplayVersion -match "15\."
        } | Sort-Object Name)
    } catch {
        try {
            $ExServers = @(Get-ClientAccessServer | Sort-Object Name)
        } catch {
            Write-Log "Cannot enumerate Exchange servers: $_" "WARN"
        }
    }

    # Build FQDN list for display
    $ExServerFqdns = @()
    foreach ($srv in $ExServers) {
        $fqdn = $srv.Fqdn
        if (-not $fqdn) { $fqdn = "$($srv.Name).$DomainFQDN" }
        $ExServerFqdns += $fqdn
    }
    $localFqdn = [System.Net.Dns]::GetHostEntry([System.Net.Dns]::GetHostName()).HostName

    if ($ExServers.Count -gt 0) {
        Write-Host ""
        Write-Host "  Available Exchange servers:" -ForegroundColor Cyan
        for ($si = 0; $si -lt $ExServers.Count; $si++) {
            $roles = $ExServers[$si].ServerRole
            Write-Host "    [$($si + 1)] $($ExServerFqdns[$si])  ($roles)" -ForegroundColor White
        }
        Write-Host ""
    } else {
        Write-Log "No Exchange servers found via cmdlets, using local FQDN: $localFqdn" "WARN"
    }

    # ---- Select SMTP server ----
    if ([string]::IsNullOrEmpty($SmtpServer)) {
        if ($ExServers.Count -eq 0) {
            $SmtpServer = $localFqdn
        } elseif ($ExServers.Count -eq 1) {
            $SmtpServer = $ExServerFqdns[0]
            Write-Log "Single Exchange server — using as SMTP: $SmtpServer" "OK"
        } else {
            $choice = Read-Host "  Select SMTP server number (1-$($ExServers.Count))"
            $choiceIdx = [int]$choice - 1
            if ($choiceIdx -lt 0 -or $choiceIdx -ge $ExServers.Count) {
                Write-Log "Invalid choice, using first server" "WARN"
                $choiceIdx = 0
            }
            $SmtpServer = $ExServerFqdns[$choiceIdx]
            Write-Log "Selected SMTP server: $SmtpServer" "OK"
        }
    } else {
        Write-Log "Using specified SMTP server: $SmtpServer"
    }

    # ---- Select EWS server ----
    if ([string]::IsNullOrEmpty($EwsServer)) {
        if ($ExServers.Count -eq 0) {
            $EwsServer = $localFqdn
        } elseif ($ExServers.Count -eq 1) {
            $EwsServer = $ExServerFqdns[0]
            Write-Log "Single Exchange server — using as EWS: $EwsServer" "OK"
        } else {
            Write-Host ""
            $ewsChoice = Read-Host "  Select EWS server number (1-$($ExServers.Count)) [same list above]"
            $ewsIdx = [int]$ewsChoice - 1
            if ($ewsIdx -lt 0 -or $ewsIdx -ge $ExServers.Count) {
                Write-Log "Invalid choice, using first server" "WARN"
                $ewsIdx = 0
            }
            $EwsServer = $ExServerFqdns[$ewsIdx]
            Write-Log "Selected EWS server: $EwsServer" "OK"
        }
    } else {
        Write-Log "Using specified EWS server: $EwsServer"
    }

    # Store in state for Phase 4
    $State.smtpServer = $SmtpServer
    $State.smtpPort = $SmtpPort
    $State.ewsServer = $EwsServer

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
    Write-Host "  Infrastructure Summary:" -ForegroundColor Cyan
    Write-Host "    Database:    $($State.database)" -ForegroundColor White
    Write-Host "    SMTP Server: $SmtpServer`:$SmtpPort" -ForegroundColor White
    Write-Host "    EWS Server:  $EwsServer" -ForegroundColor White
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

    # ---- Provision default mailbox folders (Sent Items, Drafts, etc.) ----
    # Exchange does not create default folders until the first user login.
    # We simulate a login by making an authenticated EWS GetFolder request per user.
    if (-not $State.foldersProvisioned) {
        $EwsHost = if ($State.ewsServer) { $State.ewsServer } else { $State.smtpServer }
        $EwsUrl = "https://$EwsHost/EWS/Exchange.asmx"

        Write-Log "Provisioning default mailbox folders via EWS login ($EwsUrl)..."
        Write-Log "  This triggers Exchange to create Sent Items, Drafts, etc."

        # Trust self-signed certs (lab environment)
        try {
            Add-Type @"
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public class TrustAll {
    public static void Enable() {
        ServicePointManager.ServerCertificateValidationCallback =
            delegate { return true; };
    }
}
"@
            [TrustAll]::Enable()
        } catch {
            # Type already added — ignore
        }
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

        # EWS SOAP request — GetFolder for sentitems triggers full folder tree creation
        $SoapTemplate = @'
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Body>
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="sentitems"/>
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>
'@

        # Parallel provisioning via RunspacePool
        $ProvisionBlock = {
            param($EwsUrl, $Upn, $Password)
            try {
                # Trust self-signed certs inside runspace
                Add-Type @"
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public class TrustAllRS {
    public static void Enable() {
        ServicePointManager.ServerCertificateValidationCallback =
            delegate { return true; };
    }
}
"@ -ErrorAction SilentlyContinue
                [TrustAllRS]::Enable()
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

                $soapBody = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Body>
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="sentitems"/>
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>
"@
                $cred = New-Object System.Net.NetworkCredential($Upn, $Password)
                $webReq = [System.Net.HttpWebRequest]::Create($EwsUrl)
                $webReq.Method = "POST"
                $webReq.ContentType = "text/xml; charset=utf-8"
                $webReq.Credentials = $cred
                $webReq.Timeout = 30000

                $bytes = [System.Text.Encoding]::UTF8.GetBytes($soapBody)
                $webReq.ContentLength = $bytes.Length
                $stream = $webReq.GetRequestStream()
                $stream.Write($bytes, 0, $bytes.Length)
                $stream.Close()

                $resp = $webReq.GetResponse()
                $code = [int]$resp.StatusCode
                $resp.Close()

                return @{ UPN = $Upn; Success = $true; Status = $code; Error = "" }
            } catch {
                return @{ UPN = $Upn; Success = $false; Status = 0; Error = $_.Exception.Message }
            }
        }

        $pool = [RunspaceFactory]::CreateRunspacePool(1, $Threads)
        $pool.Open()

        $jobs = @()
        foreach ($u in $Users) {
            $ps = [PowerShell]::Create().AddScript($ProvisionBlock)
            [void]$ps.AddArgument($EwsUrl)
            [void]$ps.AddArgument($u.UPN)
            [void]$ps.AddArgument($u.Password)
            $ps.RunspacePool = $pool
            $jobs += @{ PS = $ps; Handle = $ps.BeginInvoke(); UPN = $u.UPN }
        }

        $okCount = 0
        $failCount = 0
        $idx = 0
        foreach ($job in $jobs) {
            $idx++
            try {
                $result = $job.PS.EndInvoke($job.Handle)
                if ($result -and $result[0].Success) {
                    $okCount++
                } else {
                    $failCount++
                    $errMsg = if ($result) { $result[0].Error } else { "no result" }
                    Write-Log "  FAIL: $($job.UPN) — $errMsg" "WARN"
                }
            } catch {
                $failCount++
                Write-Log "  FAIL: $($job.UPN) — $_" "WARN"
            }
            $job.PS.Dispose()

            if ($idx % 50 -eq 0) {
                Write-Log "  Provisioned $idx / $($Users.Count) mailboxes (OK: $okCount, Fail: $failCount)..."
            }
        }

        $pool.Close()
        $pool.Dispose()

        Write-Log "Folder provisioning complete: $okCount OK, $failCount failed out of $($Users.Count)" $(if ($failCount -eq 0) { "OK" } else { "WARN" })

        $State.foldersProvisioned = $true
        Save-State $State
    } else {
        Write-Log "Default folders already provisioned (skipping)"
    }

    # Build a hashtable for fast credential lookup by UPN
    $UserCredMap = @{}
    foreach ($u in $Users) { $UserCredMap[$u.UPN] = $u }

    # ---- Create contacts in each user's mailbox via EWS ----
    if (-not $State.contactsCreated) {
        $EwsHost = if ($State.ewsServer) { $State.ewsServer } else { if ($State.smtpServer) { $State.smtpServer } else { $SmtpServer } }
        $EwsUrl = "https://$EwsHost/EWS/Exchange.asmx"
        Write-Log "Creating contacts in each user's mailbox via EWS ($EwsUrl)..."

        # Each user gets 10-30 random contacts from the other mock users
        $ContactBlock = {
            param($EwsUrl, $Upn, $Password, [string]$ContactsXml)

            Add-Type @"
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public class TrustAllC {
    public static void Enable() {
        ServicePointManager.ServerCertificateValidationCallback =
            delegate { return true; };
    }
}
"@ -ErrorAction SilentlyContinue
            [TrustAllC]::Enable()
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

            $cred = New-Object System.Net.NetworkCredential($Upn, $Password)
            $created = 0
            $failed = 0
            $lastError = ""

            # ContactsXml is a "|" separated list of "FirstName;LastName;Email;Company;Phone;Title;Dept"
            foreach ($entry in $ContactsXml.Split('|')) {
                if ([string]::IsNullOrWhiteSpace($entry)) { continue }
                $fields = $entry.Split(';')
                if ($fields.Count -lt 7) { continue }

                $soapBody = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Body>
    <m:CreateItem>
      <m:SavedItemFolderId>
        <t:DistinguishedFolderId Id="contacts"/>
      </m:SavedItemFolderId>
      <m:Items>
        <t:Contact>
          <t:GivenName>$($fields[0])</t:GivenName>
          <t:Surname>$($fields[1])</t:Surname>
          <t:DisplayName>$($fields[0]) $($fields[1])</t:DisplayName>
          <t:EmailAddresses>
            <t:Entry Key="EmailAddress1">$($fields[2])</t:Entry>
          </t:EmailAddresses>
          <t:CompanyName>$($fields[3])</t:CompanyName>
          <t:PhoneNumbers>
            <t:Entry Key="BusinessPhone">$($fields[4])</t:Entry>
          </t:PhoneNumbers>
          <t:JobTitle>$($fields[5])</t:JobTitle>
          <t:Department>$($fields[6])</t:Department>
        </t:Contact>
      </m:Items>
    </m:CreateItem>
  </soap:Body>
</soap:Envelope>
"@
                # Retry up to 3 times with backoff
                $ok = $false
                for ($retry = 0; $retry -lt 3; $retry++) {
                    try {
                        $webReq = [System.Net.HttpWebRequest]::Create($EwsUrl)
                        $webReq.Method = "POST"
                        $webReq.ContentType = "text/xml; charset=utf-8"
                        $webReq.Credentials = $cred
                        $webReq.Timeout = 30000

                        $bytes = [System.Text.Encoding]::UTF8.GetBytes($soapBody)
                        $webReq.ContentLength = $bytes.Length
                        $stream = $webReq.GetRequestStream()
                        $stream.Write($bytes, 0, $bytes.Length)
                        $stream.Close()

                        $resp = $webReq.GetResponse()
                        $resp.Close()
                        $created++
                        $ok = $true
                        break
                    } catch {
                        # Read error response body for diagnostics
                        $errBody = ""
                        try {
                            $errResp = $_.Exception.InnerException.Response
                            if ($errResp) {
                                $errStream = $errResp.GetResponseStream()
                                $reader = New-Object System.IO.StreamReader($errStream)
                                $errBody = $reader.ReadToEnd()
                                $reader.Close()
                            }
                        } catch {}

                        $lastError = $_.Exception.Message
                        if ($errBody -match 'faultstring[>]([^<]+)') { $lastError += " | $($Matches[1])" }

                        # Wait before retry (1s, 3s, 5s)
                        Start-Sleep -Milliseconds (($retry + 1) * 2000)
                    }
                }
                if (-not $ok) { $failed++ }

                # Small delay between contacts to avoid EWS throttling
                Start-Sleep -Milliseconds 200
            }

            $success = ($created -gt 0)
            return @{ UPN = $Upn; Success = $success; Created = $created; Failed = $failed; Error = $lastError }
        }

        $ewsThreads = [math]::Min($Threads, 5)  # Limit EWS concurrency to avoid 500 errors
        $contactPool = [RunspaceFactory]::CreateRunspacePool(1, $ewsThreads)
        $contactPool.Open()
        Write-Log "  Using $ewsThreads EWS threads for contact creation..."
        $contactJobs = @()

        foreach ($u in $Users) {
            # Pick 10-30 random contacts from other users
            $contactCount = Get-Random -Minimum 10 -Maximum 31
            $contactUsers = $Users | Where-Object { $_.UPN -ne $u.UPN } | Get-Random -Count ([math]::Min($contactCount, $Users.Count - 1))

            # Build serialized contact data string (avoid passing complex objects to runspace)
            $contactEntries = @()
            foreach ($cu in $contactUsers) {
                # DisplayName format: "MockOrg LastName FirstName"
                $nameParts = $cu.DisplayName -split ' '
                $fn = if ($cu.FirstName) { $cu.FirstName } elseif ($nameParts.Count -ge 3) { $nameParts[2] } else { $nameParts[-1] }
                $ln = if ($cu.LastName) { $cu.LastName } elseif ($nameParts.Count -ge 2) { $nameParts[1] } else { "User" }
                $comp = if ($cu.Company) { $cu.Company } else { "Global Solutions Ltd" }
                $phone = if ($cu.Phone) { $cu.Phone } else { "+1-555-$(Get-Random -Minimum 1000 -Maximum 9999)" }
                $title = if ($cu.Title) { $cu.Title } else { "Specialist" }
                $dept = if ($cu.Department) { $cu.Department } else { "IT" }
                # Escape XML special chars
                $fn = $fn -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
                $ln = $ln -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
                $comp = $comp -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
                $title = $title -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
                $dept = $dept -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
                $contactEntries += "$fn;$ln;$($cu.UPN);$comp;$phone;$title;$dept"
            }
            $contactsXml = $contactEntries -join '|'

            $ps = [PowerShell]::Create().AddScript($ContactBlock)
            [void]$ps.AddArgument($EwsUrl)
            [void]$ps.AddArgument($u.UPN)
            [void]$ps.AddArgument($u.Password)
            [void]$ps.AddArgument($contactsXml)
            $ps.RunspacePool = $contactPool
            $contactJobs += @{ PS = $ps; Handle = $ps.BeginInvoke(); UPN = $u.UPN }
        }

        $cOk = 0; $cFail = 0; $totalContacts = 0; $totalFailed = 0; $cIdx = 0
        foreach ($job in $contactJobs) {
            $cIdx++
            try {
                $result = $job.PS.EndInvoke($job.Handle)
                if ($result -and $result[0].Success) {
                    $cOk++
                    $totalContacts += $result[0].Created
                    $totalFailed += $result[0].Failed
                    if ($result[0].Failed -gt 0) {
                        Write-Log "  Partial: $($job.UPN) — $($result[0].Created) ok, $($result[0].Failed) failed" "WARN"
                    }
                } else {
                    $cFail++
                    $errMsg = if ($result) { $result[0].Error } else { "no result" }
                    Write-Log "  Contacts FAIL: $($job.UPN) — $errMsg" "WARN"
                }
            } catch {
                $cFail++
                Write-Log "  Contacts FAIL: $($job.UPN) — $_" "WARN"
            }
            $job.PS.Dispose()
            if ($cIdx % 50 -eq 0) {
                Write-Log "  Contacts: $cIdx / $($Users.Count) users processed ($totalContacts created, $totalFailed failed)..."
            }
        }

        $contactPool.Close()
        $contactPool.Dispose()
        Write-Log "Contacts created: $totalContacts across $cOk users ($cFail failed)" $(if ($cFail -eq 0) { "OK" } else { "WARN" })

        $State.contactsCreated = $true
        Save-State $State
    } else {
        Write-Log "Contacts already created (skipping)"
    }

    # ---- Send calendar meeting invitations via SMTP (iCalendar) ----
    if (-not $State.calendarCreated) {
        $SmtpServerForCal = if ($State.smtpServer) { $State.smtpServer } else { $SmtpServer }
        $CalSmtpPort = if ($State.smtpPort) { [int]$State.smtpPort } else { $SmtpPort }
        if (-not $script:TargetDC) { $script:TargetDC = (Get-ADDomain).PDCEmulator }
        $DC = $script:TargetDC
        $CalDomain = if ($DomainFQDN) { $DomainFQDN } else { (Get-ADDomain -Server $DC).DNSRoot }

        Write-Log "Sending calendar meeting invitations via SMTP ($SmtpServerForCal`:$CalSmtpPort)..."

        # Meeting subjects and locations
        $MeetingSubjects = @(
            "Weekly Team Standup", "Project Review", "Sprint Planning",
            "Budget Discussion", "Client Call", "1:1 Check-in",
            "Architecture Review", "Code Review Session", "Design Workshop",
            "All Hands Meeting", "Training Session", "Strategy Planning",
            "Incident Retrospective", "Performance Review", "Onboarding Session",
            "Product Demo", "Technical Deep Dive", "Quarterly Business Review",
            "Security Briefing", "Compliance Training", "Vendor Meeting",
            "Brainstorming Session", "Lunch & Learn", "Town Hall",
            "Совещание по проекту", "Планирование спринта", "Обзор архитектуры",
            "Reunion hebdomadaire", "Revue de projet", "Session de formation",
            "Reuniao semanal", "Revisao de projeto", "Sessao de treinamento"
        )

        $MeetingLocations = @(
            "Conference Room A", "Conference Room B", "Conference Room C",
            "Board Room", "Meeting Room 101", "Meeting Room 205",
            "Training Room", "Auditorium", "Cafeteria",
            "Microsoft Teams", "Zoom Call", "Google Meet",
            "Building A, Floor 2", "Building B, Floor 3", "Main Office",
            "Remote / Online", "HQ Large Conference Room"
        )

        $CalendarBlock = {
            param($SmtpServer, $SmtpPort, $From, $FromPassword, $FromName,
                  [string[]]$Attendees, $Subject, $Location, $StartTime, $EndTime, $Domain, $Uid)

            try {
                $dtStart = [datetime]::Parse($StartTime).ToUniversalTime().ToString("yyyyMMddTHHmmssZ")
                $dtEnd = [datetime]::Parse($EndTime).ToUniversalTime().ToString("yyyyMMddTHHmmssZ")
                $dtStamp = [datetime]::UtcNow.ToString("yyyyMMddTHHmmssZ")

                $attendeeLines = ""
                foreach ($a in $Attendees) {
                    $attendeeLines += "ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:$a`r`n"
                }

                $icsContent = @"
BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//MockData//Exchange//EN
METHOD:REQUEST
BEGIN:VEVENT
UID:$Uid
DTSTAMP:$dtStamp
DTSTART:$dtStart
DTEND:$dtEnd
SUMMARY:$Subject
LOCATION:$Location
ORGANIZER;CN=$($FromName):mailto:$From
$($attendeeLines.TrimEnd())
DESCRIPTION:This is a scheduled meeting. Please join on time.
STATUS:CONFIRMED
SEQUENCE:0
TRANSP:OPAQUE
BEGIN:VALARM
TRIGGER:-PT15M
ACTION:DISPLAY
DESCRIPTION:Reminder
END:VALARM
END:VEVENT
END:VCALENDAR
"@

                $mail = New-Object System.Net.Mail.MailMessage
                $mail.From = New-Object System.Net.Mail.MailAddress($From, $FromName)
                foreach ($addr in $Attendees) { $mail.To.Add($addr) }
                $mail.Subject = $Subject
                $mail.SubjectEncoding = [System.Text.Encoding]::UTF8

                # Create iCalendar alternate view (this makes Exchange process it as a meeting)
                $calType = New-Object System.Net.Mime.ContentType("text/calendar")
                $calType.Parameters.Add("method", "REQUEST")
                $calType.CharSet = "utf-8"
                $calView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($icsContent, $calType)
                $mail.AlternateViews.Add($calView)

                # Also add a plain text body for non-Exchange clients
                $mail.Body = "You have been invited to: $Subject`nLocation: $Location`nTime: $StartTime — $EndTime"
                $mail.IsBodyHtml = $false

                $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
                $smtp.EnableSsl = $true
                $smtp.Credentials = New-Object System.Net.NetworkCredential($From, $FromPassword)
                $smtp.Send($mail)
                $mail.Dispose()
                $smtp.Dispose()

                return @{ Success = $true; Error = $null }
            } catch {
                return @{ Success = $false; Error = $_.Exception.Message }
            }
        }

        # Each user creates 5-15 meetings with 2-8 attendees over the next 90 days
        $calPool = [RunspaceFactory]::CreateRunspacePool(1, $Threads)
        $calPool.Open()
        $calJobs = @()
        $totalMeetings = 0

        foreach ($u in $Users) {
            $meetingCount = Get-Random -Minimum 5 -Maximum 16
            for ($mi = 0; $mi -lt $meetingCount; $mi++) {
                $attendeeCount = Get-Random -Minimum 2 -Maximum 9
                $attendees = @(($Users | Where-Object { $_.UPN -ne $u.UPN } | Get-Random -Count ([math]::Min($attendeeCount, $Users.Count - 1))) | ForEach-Object { $_.UPN })

                $daysAhead = Get-Random -Minimum -30 -Maximum 91
                $hour = Get-Random -Minimum 8 -Maximum 18
                $minute = @(0, 15, 30, 45) | Get-Random
                $duration = @(30, 45, 60, 90, 120) | Get-Random
                $startDt = (Get-Date).Date.AddDays($daysAhead).AddHours($hour).AddMinutes($minute)
                $endDt = $startDt.AddMinutes($duration)
                $uid = [guid]::NewGuid().ToString() + "@$CalDomain"

                $ps = [PowerShell]::Create().AddScript($CalendarBlock)
                [void]$ps.AddArgument($SmtpServerForCal)
                [void]$ps.AddArgument($CalSmtpPort)
                [void]$ps.AddArgument($u.UPN)
                [void]$ps.AddArgument($u.Password)
                [void]$ps.AddArgument($u.DisplayName)
                [void]$ps.AddArgument($attendees)
                [void]$ps.AddArgument(($MeetingSubjects | Get-Random))
                [void]$ps.AddArgument(($MeetingLocations | Get-Random))
                [void]$ps.AddArgument($startDt.ToString("o"))
                [void]$ps.AddArgument($endDt.ToString("o"))
                [void]$ps.AddArgument($CalDomain)
                [void]$ps.AddArgument($uid)
                $ps.RunspacePool = $calPool
                $calJobs += @{ PS = $ps; Handle = $ps.BeginInvoke() }
                $totalMeetings++
            }
        }

        Write-Log "  Dispatched $totalMeetings meeting invitations across $($Users.Count) organizers..."

        $calOk = 0; $calFail = 0; $calIdx = 0
        foreach ($job in $calJobs) {
            $calIdx++
            try {
                $result = $job.PS.EndInvoke($job.Handle)
                if ($result -and $result[0].Success) { $calOk++ }
                else {
                    $calFail++
                    if ($calFail -le 5) {
                        $errMsg = if ($result) { $result[0].Error } else { "no result" }
                        Write-Log "  Calendar FAIL: $errMsg" "WARN"
                    }
                }
            } catch {
                $calFail++
            }
            $job.PS.Dispose()
            if ($calIdx % 200 -eq 0) {
                Write-Log "  Calendar: $calIdx / $totalMeetings processed (OK: $calOk, Fail: $calFail)..."
            }
        }

        $calPool.Close()
        $calPool.Dispose()
        Write-Log "Calendar events complete: $calOk sent, $calFail failed out of $totalMeetings" $(if ($calFail -eq 0) { "OK" } else { "WARN" })

        $State.calendarCreated = $true
        Save-State $State
    } else {
        Write-Log "Calendar events already created (skipping)"
    }

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
            [string[]]$To, [string[]]$Cc, [string]$Subject, [string]$HtmlBody,
            [string[]]$AttachmentPaths, [string]$InlineImagePath, [string]$InlineCid,
            [string]$InReplyTo, [string]$References
        )

        try {
            $mail = New-Object System.Net.Mail.MailMessage
            $mail.From = New-Object System.Net.Mail.MailAddress($From)
            foreach ($addr in $To) { $mail.To.Add($addr) }
            foreach ($addr in $Cc) { if ($addr) { $mail.CC.Add($addr) } }
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
            [void]$ps.AddParameter("Cc", $(if ($wi.Cc) { $wi.Cc } else { @() }))
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
            $recipientCount = Get-Random -Minimum 1 -Maximum 6  # 1-5 To recipients
            $recipients = Get-RandomRecipients -SenderUPN $sender.UPN -Count $recipientCount

            # 40% chance of CC recipients (1-4 people)
            $ccRecipients = @()
            if ((Get-Random -Minimum 1 -Maximum 101) -le 40) {
                $ccCount = Get-Random -Minimum 1 -Maximum 5
                $excludeList = @($sender.UPN) + @($recipients | ForEach-Object { $_.UPN })
                $ccPool = $Users | Where-Object { $_.UPN -notin $excludeList }
                if ($ccPool.Count -gt 0) {
                    $ccRecipients = @($ccPool | Get-Random -Count ([math]::Min($ccCount, $ccPool.Count)))
                }
            }

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
                Cc              = @($ccRecipients | ForEach-Object { $_.UPN })
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

            # 30% chance of CC'ing 1-3 other people on the reply
            $replyCc = @()
            if ((Get-Random -Minimum 1 -Maximum 101) -le 30) {
                $ccCount = Get-Random -Minimum 1 -Maximum 4
                $excludeList = @($replySenderUPN, $origMsg.SenderUPN)
                $ccPool = $Users | Where-Object { $_.UPN -notin $excludeList }
                if ($ccPool.Count -gt 0) {
                    $replyCc = @($ccPool | Get-Random -Count ([math]::Min($ccCount, $ccPool.Count)) | ForEach-Object { $_.UPN })
                }
            }

            [void]$chunk.Add(@{
                From            = $replySenderUPN
                FromPassword    = $replySenderCred.Password
                To              = @($origMsg.SenderUPN)
                Cc              = $replyCc
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

            # Forward to 1-3 To recipients + optional 1-3 CC
            $fwdToCount = Get-Random -Minimum 1 -Maximum 4
            $fwdRecipients = Get-RandomRecipients -SenderUPN $fwdSenderUPN -Count $fwdToCount

            $fwdCc = @()
            if ((Get-Random -Minimum 1 -Maximum 101) -le 35) {
                $ccCount = Get-Random -Minimum 1 -Maximum 4
                $excludeList = @($fwdSenderUPN) + @($fwdRecipients | ForEach-Object { $_.UPN })
                $ccPool = $Users | Where-Object { $_.UPN -notin $excludeList }
                if ($ccPool.Count -gt 0) {
                    $fwdCc = @($ccPool | Get-Random -Count ([math]::Min($ccCount, $ccPool.Count)) | ForEach-Object { $_.UPN })
                }
            }

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
                To              = @($fwdRecipients | ForEach-Object { $_.UPN })
                Cc              = $fwdCc
                Subject         = "FW: $($origMsg.Subject)"
                HtmlBody        = $fwdHtml
                AttachmentPaths = $attachPaths
                InlineImagePath = $null
                InlineCid       = $null
                InReplyTo       = $null
                References      = $null
                SenderName      = $origMsg.RecipientName
                RecipientUPN    = $fwdRecipients[0].UPN
                RecipientName   = $fwdRecipients[0].DisplayName
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
                $recipients = Get-RandomRecipients -SenderUPN $sender.UPN -Count (Get-Random -Minimum 1 -Maximum 6)

                # 40% chance of CC
                $extraCc = @()
                if ((Get-Random -Minimum 1 -Maximum 101) -le 40) {
                    $ccCount = Get-Random -Minimum 1 -Maximum 5
                    $excludeList = @($sender.UPN) + @($recipients | ForEach-Object { $_.UPN })
                    $ccPool = $Users | Where-Object { $_.UPN -notin $excludeList }
                    if ($ccPool.Count -gt 0) {
                        $extraCc = @($ccPool | Get-Random -Count ([math]::Min($ccCount, $ccPool.Count)) | ForEach-Object { $_.UPN })
                    }
                }

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
                    Cc = $extraCc
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
