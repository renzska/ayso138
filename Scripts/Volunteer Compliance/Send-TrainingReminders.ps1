<#
.SYNOPSIS
    Sends personalized compliance reminder emails to AYSO Region 138 volunteers,
    with role-based filtering and per-role configurable requirements.

.DESCRIPTION
    Joins two exports - the Credentials/Status file and the Volunteer Details file -
    to build a full picture of each volunteer's role, team, and compliance status.
    Sends one consolidated email per person (even if they hold multiple roles),
    listing exactly what they have completed and what they still need.

    ROLE PRESETS (default requirements per role)
    ─────────────────────────────────────────────
    Head Coach      : SafeHaven, Concussion, SCA, SafeSport*, Risk*
    Assistant Coach : SafeHaven, Concussion, SCA, SafeSport*, Risk*
    Team Manager    : SafeHaven, Concussion, SCA
    Youth Referee   : SafeHaven, Concussion, SCA       (typically under 18)
    Referee         : SafeHaven, Concussion, SCA, SafeSport*, Risk*
    Field Setup     : SafeHaven, Concussion, SCA
    Board Member    : SafeHaven, Concussion, SCA, SafeSport*, Risk*
    (* items automatically skipped for volunteers under 18)

    Use -Requirements to override presets with a custom list applied to
    all selected roles.

.PARAMETER CredentialsPath
    Path to AdminCredentialsStatusDynamic.xlsx (Sports Connect export).
    Default: "AdminCredentialsStatusDynamic.xlsx"

.PARAMETER VolunteerPath
    Path to Volunteer_Details.xlsx (Sports Connect export).
    Default: "Volunteer_Details.xlsx"

.PARAMETER GmailFrom
    Gmail address to send from.  Default: ayso138.ca@gmail.com

.PARAMETER GmailPassword
    Gmail app password (16-char, not your regular password).

.PARAMETER Roles
    Comma-separated list of roles to process, or "All".
    Accepted values (case-insensitive, spaces optional):
      HeadCoach (HC) | AssistantCoach (AC) | TeamManager (TM) |
      YouthReferee (YR) | Referee (Ref) | FieldSetup (FS) |
      BoardMember (BM) | All
    Examples:
      -Roles "HeadCoach,AssistantCoach"
      -Roles "All"
      -Roles "TM"

.PARAMETER Requirements
    Override the role presets for ALL selected roles.
    Comma-separated from: SafeHaven (SH), Concussion (CDC), SCA,
    SafeSport (SS), Risk (BGC) - or the word "All".
    Leave empty (default) to use role presets.
    Examples:
      -Requirements "SafeHaven,Concussion,SCA,SafeSport,Risk"   ← all five
      -Requirements "Risk"                                        ← background check only
      -Requirements "All"                                         ← same as all five

.PARAMETER TestMode
    Sends all emails to GmailFrom instead of the real volunteer.
    Use this to review every email before going live.

.PARAMETER DryRun
    No emails sent. Saves HTML previews and prints a full summary.

.PARAMETER ExportReport
    Writes a CSV report (ComplianceReport.csv) with every volunteer's
    status - useful for reviewing or sharing with the RC.

.PARAMETER Division
    Partial, case-insensitive filter on division name.
    Examples: -Division "12U"  -Division "6U"  -Division "Playground"

.PARAMETER CC
    If provided, CC this address on every email sent.

.PARAMETER SkipUnallocated
    Skip volunteers whose only team assignment is "Unallocated".

.PARAMETER OutputFolder
    Folder for HTML email previews.  Default: "TrainingReminderOutput"

.PARAMETER OnlyList
    Path to a plain-text file of email addresses (one per line) to send to exclusively.
    Useful for retrying a handful of failed sends without reprocessing everyone.
    Example: -OnlyList "TrainingReminderOutput\FailedEmails.txt"

.PARAMETER SentLogPath
    Path to a plain-text file where successfully sent addresses are recorded (one per line).
    Default: "TrainingReminderOutput\SentLog.txt"
    Used together with -SkipAlreadySent to avoid double-sending on re-runs.

.PARAMETER SkipAlreadySent
    When set, any email address already present in SentLogPath is skipped.
    Use this on a full re-run after a partial failure so already-sent volunteers
    are not emailed again.

.PARAMETER DelayMs
    Milliseconds to pause between each send attempt. Helps avoid Gmail SMTP rate
    limiting. Default: 1500 ms (safe for large lists). Lower to 500 for small retries.

.PARAMETER SafetyOnly
    When set, only the five safety-training items are evaluated and included in emails:
      SafeHaven, Concussion (CDC), SCA, SafeSport, Risk (Background Check)
    Coach training, referee training, and lower-division course checks are completely
    skipped — they will not appear in the status table, email instructions, or
    compliance determination. Use this for a clean safety-focused pass before or
    after running a separate coach-training pass.

.PARAMETER SendCompliantEmail
    When set, any volunteer who has completed all required safety training receives
    a separate congratulatory "You're All Set" email confirming their safety compliance.
    This email is sent instead of a reminder — compliant volunteers will not receive
    both. Pair with -SafetyOnly for the most common use case:
      - Reminders  → go to volunteers still missing safety items
      - Congrats   → go to volunteers who have finished all safety items
    Can also be used without -SafetyOnly; in that case, a volunteer who is
    safety-complete but still needs coach training will receive the congrats email
    AND a separate reminder about the outstanding coach training.

.EXAMPLE
    # Preview everything — no emails sent, HTML previews saved for review
    .\Send-TrainingReminders.ps1 -DryRun -ExportReport

.EXAMPLE
    # Test mode: all emails land in your own inbox so you can review before going live
    .\Send-TrainingReminders.ps1 -Roles "HeadCoach,AssistantCoach" -TestMode

.EXAMPLE
    # Send only to Team Managers (uses role preset: SafeHaven, Concussion, SCA)
    .\Send-TrainingReminders.ps1 -Roles "TeamManager"

.EXAMPLE
    # Send to Team Managers + Field Setup, but require ALL five training items
    .\Send-TrainingReminders.ps1 -Roles "TM,FS" -Requirements "All"

.EXAMPLE
    # 12U coaches only, test mode
    .\Send-TrainingReminders.ps1 -Roles "HC,AC" -Division "12U" -TestMode

.EXAMPLE
    # Live send to everyone, CC yourself
    .\Send-TrainingReminders.ps1 -Roles All -CC "john@rennemeyer.com"

.EXAMPLE
    # Preview safety-only reminders — coach and referee training completely excluded
    .\Send-TrainingReminders.ps1 -Roles All -SafetyOnly -DryRun

.EXAMPLE
    # Test safety-only reminders in your inbox before going live
    .\Send-TrainingReminders.ps1 -Roles All -SafetyOnly -TestMode

.EXAMPLE
    # Live safety-only reminders to everyone still missing safety training
    .\Send-TrainingReminders.ps1 -Roles All -SafetyOnly -CC "john@rennemeyer.com"

.EXAMPLE
    # Preview both safety reminders AND congrats emails — nothing sent yet
    .\Send-TrainingReminders.ps1 -Roles All -SafetyOnly -SendCompliantEmail -DryRun

.EXAMPLE
    # Test both in your inbox: reminders to those missing safety, congrats to those done
    .\Send-TrainingReminders.ps1 -Roles All -SafetyOnly -SendCompliantEmail -TestMode

.EXAMPLE
    # Live full safety pass: reminders to the incomplete, congrats to the complete
    .\Send-TrainingReminders.ps1 -Roles All -SafetyOnly -SendCompliantEmail -CC "john@rennemeyer.com"

.EXAMPLE
    # Preview congrats-only emails (safety complete check, no reminders sent)
    .\Send-TrainingReminders.ps1 -Roles All -SendCompliantEmail -DryRun
#>

param (
    [string]$CredentialsPath  = "AdminCredentialsStatusDynamic.xlsx",
    [string]$VolunteerPath    = "Volunteer_Details.xlsx",
    [string]$GmailFrom        = "ayso138.ca@gmail.com",
    [string]$GmailPassword    = "rovn yjfh ansj popl",
    [string]$Roles            = "All",
    [string]$Requirements     = "",
    [switch]$TestMode         = $false,
    [switch]$DryRun           = $false,
    [switch]$ExportReport     = $false,
    [string]$Division         = "",
    [string]$CC               = "",
    [switch]$SkipUnallocated  = $false,
    [string]$OutputFolder     = "TrainingReminderOutput",

    # ── Retry / throttle controls ──────────────────────────────────────────
    # Path to a plain-text file of email addresses (one per line) to send to exclusively.
    # Useful for retrying failed sends: point this at FailedEmails.txt.
    [string]$OnlyList         = "",

    # Path to the sent-log file. After each successful send, the recipient's
    # email address is appended here so future runs can skip them.
    [string]$SentLogPath      = "TrainingReminderOutput\SentLog.txt",

    # When set, any email address already present in SentLogPath is skipped.
    # Use this on a full re-run to avoid double-sending to people who already received the email.
    [switch]$SkipAlreadySent  = $false,

    # Milliseconds to pause between each send attempt. Helps avoid Gmail SMTP
    # rate limiting. 1500 ms is safe for large lists; lower to 500 for small retries.
    [int]$DelayMs             = 1500,

    # When set, only safety-training items are evaluated (SafeHaven, Concussion, SCA,
    # SafeSport, Risk). Coach training, referee training, and lower-division checks are
    # completely skipped — they will not appear in emails or compliance checks.
    [switch]$SafetyOnly         = $false,

    # When set, volunteers who have completed all required safety training receive a
    # congratulatory email. Pair with -SafetyOnly for a clean safety-only pass.
    [switch]$SendCompliantEmail = $false
)

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

# ============================================================================
# MODULE CHECK
# ============================================================================

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "📦 ImportExcel not found - installing..." -ForegroundColor Yellow
    Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
}
Import-Module ImportExcel -ErrorAction Stop

# ============================================================================
# ROLE NAME NORMALIZER
# Maps friendly/short names → exact strings used in the volunteer file
# ============================================================================

function Resolve-RoleName ([string]$raw) {
    $map = @{
        'headcoach'      = 'Head Coach'
        'hc'             = 'Head Coach'
        'head'           = 'Head Coach'
        'assistantcoach' = 'Assistant Coach'
        'assistant'      = 'Assistant Coach'
        'ac'             = 'Assistant Coach'
        'teammanager'    = 'Team Manager'
        'manager'        = 'Team Manager'
        'tm'             = 'Team Manager'
        'youthreferee'   = 'Youth Referee'
        'yr'             = 'Youth Referee'
        'youth'          = 'Youth Referee'
        'referee'        = 'Referee'
        'ref'            = 'Referee'
        'fieldsetup'     = 'Field Setup'
        'setup'          = 'Field Setup'
        'field'          = 'Field Setup'
        'fs'             = 'Field Setup'
        'boardmember'    = 'Board Member'
        'board'          = 'Board Member'
        'bm'             = 'Board Member'
        'all'            = 'All'
    }
    $key = $raw.ToLower().Trim() -replace '\s+', ''
    if ($map.ContainsKey($key)) { return $map[$key] }
    return $null
}

# ============================================================================
# REQUIREMENT PARSER
# Returns $null (use presets) or an array of requirement keys
# ============================================================================

function Parse-Requirements ([string]$raw) {
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }

    $reqMap = @{
        'safehaven'       = 'SafeHaven'
        'sh'              = 'SafeHaven'
        'haven'           = 'SafeHaven'
        'concussion'      = 'Concussion'
        'cdc'             = 'Concussion'
        'concuss'         = 'Concussion'
        'sca'             = 'SCA'
        'cardiac'         = 'SCA'
        'suddencardiac'   = 'SCA'
        'safesport'       = 'SafeSport'
        'ss'              = 'SafeSport'
        'sport'           = 'SafeSport'
        'risk'            = 'Risk'
        'bgc'             = 'Risk'
        'background'      = 'Risk'
        'backgroundcheck' = 'Risk'
    }
    $allItems = @('SafeHaven','Concussion','SCA','SafeSport','Risk')

    if ($raw.Trim().ToLower() -eq 'all') { return $allItems }

    $result = @()
    foreach ($item in ($raw -split ',')) {
        $key = $item.ToLower().Trim() -replace '\s+', ''
        if ($reqMap.ContainsKey($key)) {
            $result += $reqMap[$key]
        } else {
            Write-Warning "⚠️  Unknown requirement '$item' - valid: SafeHaven, Concussion, SCA, SafeSport, Risk, All"
        }
    }
    return ($result | Select-Object -Unique)
}

# ============================================================================
# ROLE PRESETS
# Default requirements for each role.  SafeSport and Risk are still subject
# to the 18+ age filter at runtime regardless of what's listed here.
# ============================================================================

$RolePresets = @{
    'Head Coach'      = @('SafeHaven','Concussion','SCA','SafeSport','Risk')
    'Assistant Coach' = @('SafeHaven','Concussion','SCA','SafeSport','Risk')
    'Team Manager'    = @('SafeHaven','Concussion','SCA')
    'Youth Referee'   = @('SafeHaven','Concussion','SCA')
    'Referee'         = @('SafeHaven','Concussion','SCA','SafeSport','Risk')
    'Field Setup'     = @('SafeHaven','Concussion','SCA')
    'Board Member'    = @('SafeHaven','Concussion','SCA','SafeSport','Risk')
}

# ============================================================================
# COACH TRAINING LEVEL HIERARCHY
# Source column: "Coaching License Level" in AdminCredentialsStatusDynamic.xlsx
# Higher rank = higher certification. A higher-ranked coach can coach any
# lower-ranked division.
# ============================================================================

$CoachLevelRank = @{
    ''                                      = 0
    'none'                                  = 0
    'playground/schoolyard activity leader' = 0   # old label, treated as no license
    'kickstart soccer play leader'          = 1
    '6u coach'                              = 2
    '8u coach'                              = 3
    '10u coach'                             = 4
    '12u coach'                             = 5
    'intermediate (14u) coach'              = 6
    'advanced (16u-19u) coach'              = 7
}

# Minimum rank required to coach each division age group.
$DivisionRequiredRank = @{
    0  = 1   # Playground / Schoolyard / Kickstart → Kickstart Soccer Play Leader
    6  = 2   # 6U  → 6U Coach
    8  = 3   # 8U  → 8U Coach
    10 = 4   # 10U → 10U Coach
    12 = 5   # 12U → 12U Coach
    13 = 6   # JH / HS → Intermediate (14U); path is 12U online then 14U online
    99 = 6
}

# Friendly label for the required level, used in the email body
$DivisionRequiredLabel = @{
    0  = 'PG/SY/KS Activity Leader'
    6  = '6U Coach'
    8  = '8U Coach'
    10 = '10U Coach'
    12 = '12U Coach'
    13 = 'Intermediate (14U) Coach'
    99 = 'Intermediate (14U) Coach'
}

# AYSOU enrollment text shown inside the numbered steps of the email
$DivisionCourseLabel = @{
    0  = 'Kickstart Soccer Play Leader - Full Online Course'
    6  = '6U Coach - Full Online Course'
    8  = '8U Coach - Full Online Course'
    10 = '10U Coach - Online + In-Person Course'
    12 = '12U Coach - Online + In-Person Course'
    13 = 'see steps below'   # JH/HS has a two-step path handled separately
}

# ── Helper: extract the numeric age from a division name ─────────────────────
# "Spring 12U - Co-Ed" → 12,  "Spring Jr. High/High School" → 13,  "Spring Playground" → 0
function Get-DivisionAge ([string]$divName) {
    if ($divName -imatch 'jr.*high|high.?school|transition|13u|14u|15u|16u|17u|18u|19u') { return 13 }
    if ($divName -imatch 'playground|schoolyard|kickstart') { return 0 }
    if ($divName -imatch '(\d+)U') { return [int]$Matches[1] }
    return -1   # unknown / unrecognized
}

# ── Helper: map a division name → short filename prefix label ─────────────────
# Distinguishes Playground/Schoolyard/Kickstart unlike Get-DivisionAge.
function Get-DivisionLabel ([string]$divName) {
    if ($divName -imatch 'jr.*high|high.?school|transition|13u|14u|15u|16u|17u|18u|19u') { return 'JHHS' }
    if ($divName -imatch 'playground')  { return 'PG' }
    if ($divName -imatch 'schoolyard')  { return 'SY' }
    if ($divName -imatch 'kickstart')   { return 'KS' }
    if ($divName -imatch '(\d+)U')      { return "$($Matches[1])U" }
    return ''
}

# ============================================================================
# RESOLVE -Roles PARAMETER
# ============================================================================

$allPossibleRoles = @('Head Coach','Assistant Coach','Team Manager',
                      'Youth Referee','Referee','Field Setup','Board Member')

$selectedRoles = @()
if ($Roles.Trim().ToLower() -eq 'all') {
    $selectedRoles = $allPossibleRoles
} else {
    foreach ($r in ($Roles -split ',')) {
        $resolved = Resolve-RoleName $r.Trim()
        if ($resolved -and $resolved -ne 'All') {
            $selectedRoles += $resolved
        } elseif ($resolved -eq 'All') {
            $selectedRoles = $allPossibleRoles
            break
        } else {
            Write-Warning "⚠️  Unknown role '$($r.Trim())' - skipping. Valid: HeadCoach, AssistantCoach, TeamManager, YouthReferee, Referee, FieldSetup, BoardMember, All"
        }
    }
    $selectedRoles = $selectedRoles | Select-Object -Unique
}

if ($selectedRoles.Count -eq 0) {
    Write-Error "❌ No valid roles resolved from -Roles '$Roles'. Exiting."
    exit 1
}

# ============================================================================
# RESOLVE -Requirements PARAMETER
# ============================================================================

$overrideRequirements = Parse-Requirements $Requirements
# $null means "use presets"; @() or array means "use this list for everyone"

# ============================================================================
# SETUP OUTPUT FOLDER
# ============================================================================

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

# ============================================================================
# MODE BANNER
# ============================================================================

# ============================================================================
# LOAD ONLY-LIST  (optional - send exclusively to these addresses)
# ============================================================================

$onlySet = $null
if ($OnlyList) {
    if (-not (Test-Path $OnlyList)) {
        Write-Error "❌ OnlyList file not found: $OnlyList"
        exit 1
    }
    $onlySet = @{}
    Get-Content $OnlyList | ForEach-Object {
        $addr = $_.Trim().ToLower()
        if ($addr -and $addr -match "@") { $onlySet[$addr] = $true }
    }
    Write-Host "📋 OnlyList loaded: $($onlySet.Count) address(es) from $OnlyList" -ForegroundColor Yellow
}

# ============================================================================
# LOAD SENT LOG  (optional - skip already-sent addresses)
# ============================================================================

$sentSet = @{}
if ($SkipAlreadySent -and (Test-Path $SentLogPath)) {
    Get-Content $SentLogPath | ForEach-Object {
        $addr = $_.Trim().ToLower()
        if ($addr -and $addr -match "@") { $sentSet[$addr] = $true }
    }
    Write-Host "📋 SentLog loaded: $($sentSet.Count) address(es) already sent — will skip." -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "══════════════════════════════════════════════════" -ForegroundColor DarkCyan
Write-Host "  AYSO Region 138 - Training Compliance Mailer"   -ForegroundColor White
Write-Host "══════════════════════════════════════════════════" -ForegroundColor DarkCyan

$modeLabel = if ($DryRun) { "Dry Run (no emails)" } elseif ($TestMode) { "Test (→ $GmailFrom)" } else { "LIVE" }
Write-Host "  Mode         : $modeLabel" -ForegroundColor $(if ($DryRun) {"Cyan"} elseif ($TestMode) {"Yellow"} else {"Green"})
Write-Host "  Roles        : $($selectedRoles -join ', ')"
if ($overrideRequirements -ne $null) {
    Write-Host "  Requirements : $($overrideRequirements -join ', ') (override)" -ForegroundColor Yellow
} else {
    Write-Host "  Requirements : Role presets"
}
if ($SafetyOnly)         { Write-Host "  Safety Only  : Yes (coach/ref training skipped)" -ForegroundColor Cyan }
if ($SendCompliantEmail) { Write-Host "  Compliant Emails: Yes (sending to safety-complete volunteers)" -ForegroundColor Cyan }
if ($Division) { Write-Host "  Division     : *$Division*" }
if ($CC)       { Write-Host "  CC           : $CC" }
Write-Host ""

# ============================================================================
# LOAD FILES
# ============================================================================

Write-Host "📂 Loading credentials:  $CredentialsPath" -ForegroundColor Cyan
if (-not (Test-Path $CredentialsPath)) { Write-Error "❌ Not found: $CredentialsPath"; exit 1 }
$credRows = Import-Excel -Path $CredentialsPath

Write-Host "📂 Loading volunteer details: $VolunteerPath" -ForegroundColor Cyan
if (-not (Test-Path $VolunteerPath))   { Write-Error "❌ Not found: $VolunteerPath"; exit 1 }
$volRows  = Import-Excel -Path $VolunteerPath

# ============================================================================
# BUILD CREDENTIAL LOOKUP  (email.lower → credential row)
# ============================================================================

$credLookup = @{}
$credNameLookup = @{}  # fallback: "firstname lastname".lower → credential row
foreach ($c in $credRows) {
    $em = "$($c.'Email')".Trim().ToLower()
    if ($em -and $em -match "@") {
        $credLookup[$em] = $c
    }
    
    # Also build name-based lookup for fallback
    $cfn = "$($c.'First Name')".Trim()
    $cln = "$($c.'Last Name')".Trim()
    if ($cfn -and $cln) {
        $nameKey = "$cfn $cln".ToLower()
        # Only store if we don't already have this name (avoid duplicates)
        if (-not $credNameLookup.ContainsKey($nameKey)) {
            $credNameLookup[$nameKey] = $c
        }
    }
}

# ============================================================================
# BUILD VOLUNTEER LOOKUP  (email.lower → list of role records)
# ============================================================================

$volLookup = @{}   # email → @( @{Role; Division; Team; FirstName; LastName; Phone} )

foreach ($v in $volRows) {
    $em   = "$($v.'Volunteer Email Address')".Trim().ToLower()
    $role = "$($v.'Volunteer Role')".Trim()
    $div  = "$($v.'Division Name')".Trim()
    $team = "$($v.'Team Name')".Trim()
    $fn   = "$($v.'Volunteer First Name')".Trim()
    $ln   = "$($v.'Volunteer Last Name')".Trim()

    # Prefer cell, fall back to telephone, then other phone
    $cell  = "$($v.'Volunteer Cellphone')".Trim()
    $tel   = "$($v.'Volunteer Telephone')".Trim()
    $other = "$($v.'Volunteer Other Phone')".Trim()
    $phone = if ($cell)  { $cell }
             elseif ($tel)   { $tel }
             elseif ($other) { $other }
             else            { '' }

    if (-not $em -or $em -notmatch "@") { continue }
    if (-not $volLookup.ContainsKey($em)) { $volLookup[$em] = @() }

    # Deduplicate: skip if exact same role+team already recorded
    $existing = $volLookup[$em] | Where-Object { $_.Role -eq $role -and $_.Team -eq $team }
    if (-not $existing) {
        $volLookup[$em] += @{ Role=$role; Division=$div; Team=$team; FirstName=$fn; LastName=$ln; Phone=$phone }
    }
}

# ============================================================================
# HELPER: HTML STATUS BADGE
# ============================================================================

function Get-Badge ([bool]$Done, [bool]$NotRequired = $false) {
    if ($NotRequired) { return '<span style="color:#888888;">N/A - not required</span>' }
    if ($Done)        { return '<span style="color:#1a7a1a;font-weight:bold;">&#10003; Complete</span>' }
    else              { return '<span style="color:#cc0000;font-weight:bold;">&#10007; Needed</span>' }
}

# ============================================================================
# REPORT ACCUMULATOR
# ============================================================================

$reportRows         = @()
$compliantList      = @()   # accumulates fully-compliant volunteers for summary
$nonCompliantList   = @()   # accumulates volunteers who received a reminder email
$warningList        = @()   # accumulates volunteers with no credentials record
$sentCount          = 0
$skippedCount       = 0
$warnCount          = 0
$errorCount         = 0
$today              = Get-Date

# ============================================================================
# MAIN LOOP - iterate unique volunteer emails
# ============================================================================

$processedEmails = @{}

foreach ($em in $volLookup.Keys) {

    $allRecords = $volLookup[$em]

    # --- Filter by selected roles ---
    $matchingRecords = $allRecords | Where-Object { $selectedRoles -contains $_.Role }
    if (-not $matchingRecords -or @($matchingRecords).Count -eq 0) { continue }

    # --- Division filter ---
    if ($Division) {
        $matchingRecords = $matchingRecords | Where-Object {
            $_.Division -imatch [regex]::Escape($Division)
        }
        if (-not $matchingRecords -or @($matchingRecords).Count -eq 0) { continue }
    }

    # --- Skip unallocated-only if requested ---
    if ($SkipUnallocated) {
        $realTeams = $matchingRecords | Where-Object { $_.Team -ne 'Unallocated' }
        if (-not $realTeams -or @($realTeams).Count -eq 0) {
            Write-Host "⏭️  Skipping (Unallocated only): $($allRecords[0].FirstName) $($allRecords[0].LastName)" -ForegroundColor DarkGray
            continue
        }
    }

    # Guard against processing the same person twice
    if ($processedEmails.ContainsKey($em)) { continue }
    $processedEmails[$em] = $true

    # --- OnlyList filter: skip anyone not in the explicit list ---
    if ($onlySet -ne $null -and -not $onlySet.ContainsKey($em)) { continue }

    # --- SentLog filter: skip anyone already successfully sent to ---
    if ($sentSet.ContainsKey($em)) {
        Write-Host "⏭️  Already sent, skipping: $(($volLookup[$em] | Select-Object -First 1).FirstName) $(($volLookup[$em] | Select-Object -First 1).LastName) ($em)" -ForegroundColor DarkGray
        continue
    }

    $firstName = ($matchingRecords | Select-Object -First 1).FirstName
    $lastName  = ($matchingRecords | Select-Object -First 1).LastName
    $displayEmail = $em   # preserve original case from cred lookup if possible

    # --- Look up credentials ---
    $cred = $null
    $noCredRecord = $false
    $matchedByName = $false
    
    # First try email lookup
    if ($credLookup.ContainsKey($em)) {
        $cred = $credLookup[$em]
        # Preserve original email from cred file for sending
        $displayEmail = "$($cred.'Email')".Trim()
    } else {
        # Email not found - try matching by name
        $nameKey = "$firstName $lastName".ToLower()
        if ($credNameLookup.ContainsKey($nameKey)) {
            $cred = $credNameLookup[$nameKey]
            $matchedByName = $true
            $credEmail = "$($cred.'Email')".Trim()
            Write-Host "ℹ️  Matched by name: $firstName $lastName - Volunteer email: $em, Cred email: $credEmail" -ForegroundColor Cyan
            # Use the email from the credentials file for sending
            $displayEmail = $credEmail
            
            # Track this as a warning for email mismatch
            $warnCount++
            $warningList += [PSCustomObject]@{
                FirstName = $firstName
                LastName  = $lastName
                Email     = "Vol: $em / Cred: $credEmail"
                Roles     = (($matchingRecords | Select-Object -ExpandProperty Role | Sort-Object -Unique) -join ', ')
                Issue     = 'Email mismatch - matched by name'
            }
        } else {
            # No match by email or name
            Write-Host "⚠️  No credentials record: $firstName $lastName ($em) - treating as non-compliant" -ForegroundColor Yellow
            $warnCount++
            $noCredRecord = $true
            
            # Track for summary
            $warningList += [PSCustomObject]@{
                FirstName = $firstName
                LastName  = $lastName
                Email     = $em
                Roles     = (($matchingRecords | Select-Object -ExpandProperty Role | Sort-Object -Unique) -join ', ')
                Issue     = 'Not in credentials file'
            }

            # Create a fake credential record with all fields indicating incomplete
            $cred = [PSCustomObject]@{
                'Email'                              = $em
                'DOB'                                = $null
                'Risk Status'                        = 'None'
                'AYSOs Safe Haven Verified'          = 'N'
                'Concussion Awareness Verified'      = 'N'
                'SafeSport Verified'                 = 'N'
                'Sudden Cardiac Arrest Verified'     = 'N'
                'Referee Grade'                      = ''
            }
            $displayEmail = $em
        }
    }

    # --- Age / adult determination ---
    $isAdult = $true
    $age     = $null
    $dobRaw  = $cred.'DOB'
    if ($dobRaw) {
        try {
            $dob     = [datetime]$dobRaw
            $age     = [math]::Floor(($today - $dob).TotalDays / 365.25)
            $isAdult = ($age -ge 18)
        } catch { }
    }

    # Youth Referee is under-18 by role definition, not just by age.
    # If every matching role for this person is an under-18 role, force
    # $isAdult = $false regardless of DOB, missing DOB, or -Requirements override.
    $under18Roles = @('Youth Referee')
    $hasAdultRole = @($matchingRecords | Where-Object { $_.Role -notin $under18Roles }).Count -gt 0
    if (-not $hasAdultRole) { $isAdult = $false }

    # --- Read credential values ---
    $riskStatus = "$($cred.'Risk Status')".Trim()
    $safeHaven  = "$($cred.'AYSOs Safe Haven Verified')".Trim()
    $concussion = "$($cred.'Concussion Awareness Verified')".Trim()
    $safeSport  = "$($cred.'SafeSport Verified')".Trim()
    $sca        = "$($cred.'Sudden Cardiac Arrest Verified')".Trim()
    $refGrade   = "$($cred.'Referee Grade')".Trim()

    # Blue Risk Status is the authoritative signal that this is a youth volunteer.
    # Strip SafeSport and Risk from required items entirely - not as Needed,
    # not as Complete, not even as N/A.
    $isBlueStatus = ($riskStatus -ieq 'Blue')
    if ($isBlueStatus) { $isAdult = $false }

    # --- Determine UNION of required items across all matching roles ---
    $requiredItems = @()
    if ($overrideRequirements -ne $null) {
        $requiredItems = $overrideRequirements
    } else {
        foreach ($rec in $matchingRecords) {
            $preset = $RolePresets[$rec.Role]
            if ($preset) { $requiredItems += $preset }
        }
        $requiredItems = $requiredItems | Select-Object -Unique
    }

    # If Blue status, remove SafeSport and Risk from the required list entirely.
    # This suppresses them from the table and from N/A rows regardless of role presets.
    if ($isBlueStatus) {
        $requiredItems = @($requiredItems | Where-Object { $_ -notin @('SafeSport','Risk') })
    }

    # --- Enforce 18+ rules: SafeSport and Risk never checked for under-18 ---
    $checkSafeHaven  = $requiredItems -contains 'SafeHaven'
    $checkConcussion = $requiredItems -contains 'Concussion'
    $checkSCA        = $requiredItems -contains 'SCA'
    $checkSafeSport  = ($requiredItems -contains 'SafeSport') -and $isAdult
    $checkRisk       = ($requiredItems -contains 'Risk') -and $isAdult

    # --- Compliance results ---
    $doneSafeHaven  = ($safeHaven  -eq 'Y')
    $doneConcussion = ($concussion -eq 'Y')
    $doneSCA        = ($sca        -eq 'Y')
    $doneSafeSport  = ($safeSport  -eq 'Y')
    $doneRisk       = ($riskStatus -in @('Green','Blue'))

    $needsSafeHaven  = $checkSafeHaven  -and (-not $doneSafeHaven)
    $needsConcussion = $checkConcussion -and (-not $doneConcussion)
    $needsSCA        = $checkSCA        -and (-not $doneSCA)
    $needsSafeSport  = $checkSafeSport  -and (-not $doneSafeSport)
    $needsRisk       = $checkRisk       -and (-not $doneRisk)

    # ── Coach Training Check (Head Coach / Assistant Coach only) ────────────────
    $needsCoachTraining  = $false
    $coachTrainingStatus = $null   # $null = N/A, $true = current, $false = needs update
    $coachHighestAge     = -1
    $coachCurrentLevel   = 'None'
    $coachCurrentRank    = 0
    $coachRequiredRank   = 0
    $coachRequiredLabel  = ''
    $coachCourseLabel    = ''
    $coachIsJHHS         = $false
    $allCoachRecs        = @()    # always initialize so filename prefix can reference it

    $isCoachRole = @($matchingRecords | Where-Object { $_.Role -in @('Head Coach', 'Assistant Coach') }).Count -gt 0
    # Always populate allCoachRecs — needed for filename generation even in SafetyOnly mode
    if ($isCoachRole) {
        $allCoachRecs = @($allRecords | Where-Object { $_.Role -in @('Head Coach', 'Assistant Coach') })
    }
    if ($isCoachRole -and -not $SafetyOnly) {
        # Use ALL records for this person (including unallocated) to find their
        # highest coaching division - that determines the training level required
        foreach ($rec in $allCoachRecs) {
            $age = Get-DivisionAge $rec.Division
            if ($age -gt $coachHighestAge) { $coachHighestAge = $age }
        }

        # Read current coaching license from credentials
        $rawLevel         = "$($cred.'Coaching License Level')".Trim()
        $coachCurrentLevel = if ($rawLevel -and $rawLevel -notin @('','None')) { $rawLevel } else { 'None' }
        $levelKey          = $rawLevel.ToLower().Trim()
        $coachCurrentRank  = if ($CoachLevelRank.ContainsKey($levelKey)) { $CoachLevelRank[$levelKey] } else { 0 }

        if ($coachHighestAge -eq 13 -or $coachHighestAge -ge 99) {
            # JH / HS: target is Intermediate (14U) rank 4, reached via 12U online then 14U online
            $coachIsJHHS        = $true
            $coachRequiredRank  = 6
            $coachRequiredLabel = 'Intermediate (14U) Coach'
            if ($coachCurrentRank -ge $coachRequiredRank) {
                $coachTrainingStatus = $true
            } else {
                $coachTrainingStatus = $false
                $needsCoachTraining  = $true
            }
        } elseif ($coachHighestAge -ge 0) {
            $coachRequiredRank  = if ($DivisionRequiredRank.ContainsKey($coachHighestAge))  { $DivisionRequiredRank[$coachHighestAge]  } else { 0 }
            $coachRequiredLabel = if ($DivisionRequiredLabel.ContainsKey($coachHighestAge)) { $DivisionRequiredLabel[$coachHighestAge] } else { '' }
            $coachCourseLabel   = if ($DivisionCourseLabel.ContainsKey($coachHighestAge))   { $DivisionCourseLabel[$coachHighestAge]   } else { '' }

            if ($coachCurrentRank -ge $coachRequiredRank) {
                $coachTrainingStatus = $true    # fully current for their division
            } else {
                $coachTrainingStatus = $false
                $needsCoachTraining  = $true
            }
        }
        # $coachHighestAge -lt 0 → unrecognized division, leave status $null
    }

    # ── Youth Referee Training Check ─────────────────────────────────────────────
    # Only check 8U Official via Referee Grade column.
    $needsRefTraining = $false
    $needs8UOfficial  = $false
    $isYouthRefRole   = @($matchingRecords | Where-Object { $_.Role -eq 'Youth Referee' }).Count -gt 0
    if ((-not $SafetyOnly) -and ($isBlueStatus -or $isYouthRefRole)) {
        $needs8UOfficial  = ($refGrade -eq '' -or $refGrade -ieq 'None')
        $needsRefTraining = $needs8UOfficial
    }

    # ── Lower-Division Coach Check ────────────────────────────────────────────
    # For each lower division a coach is actually assigned to, check whether they
    # hold the corresponding license. These are independent of the main coach
    # training check (which targets their highest division).
    $needsCoach6U        = $false
    $needsCoach8U        = $false
    $needsLowerDivVerify = $false
    if ((-not $SafetyOnly) -and $isCoachRole -and $allCoachRecs.Count -gt 0) {
        $coaches6U = @($allCoachRecs | Where-Object { (Get-DivisionAge $_.Division) -eq 6 }).Count -gt 0
        $coaches8U = @($allCoachRecs | Where-Object { (Get-DivisionAge $_.Division) -eq 8 }).Count -gt 0
        if ($coaches6U -and $coachCurrentRank -lt 2) { $needsCoach6U        = $true }
        if ($coaches8U -and $coachCurrentRank -lt 3) { $needsCoach8U        = $true }
        $needsLowerDivVerify = $needsCoach6U -or $needsCoach8U
    }

    $safetyComplete = -not ($needsSafeHaven -or $needsConcussion -or $needsSCA -or $needsSafeSport -or $needsRisk)
    $anyIncomplete  = $needsSafeHaven -or $needsConcussion -or $needsSCA -or $needsSafeSport -or $needsRisk -or $needsCoachTraining -or $needsRefTraining -or $needsLowerDivVerify

    # --- Build missing items list for report ---
    $missingList = @()
    if ($needsRisk)          { $missingList += "Risk Status" }
    if ($needsSafeHaven)     { $missingList += "Safe Haven" }
    if ($needsConcussion)    { $missingList += "Concussion" }
    if ($needsSCA)           { $missingList += "SCA" }
    if ($needsSafeSport)     { $missingList += "SafeSport" }
    if ($needsCoachTraining) { $missingList += "Coach Training" }
    if ($needsRefTraining)   { $missingList += "Referee Training" }
    if ($needsCoach6U)       { $missingList += "6U Coach Course" }
    if ($needsCoach8U)       { $missingList += "8U Coach Course" }

    # --- Roles / teams summary strings ---
    $rolesSummary  = ($matchingRecords | Select-Object -ExpandProperty Role    | Sort-Object -Unique) -join ', '
    $divsSummary   = ($matchingRecords | Select-Object -ExpandProperty Division | Sort-Object -Unique) -join ', '
    $teamsSummary  = ($matchingRecords | Select-Object -ExpandProperty Team     | Sort-Object -Unique) -join ', '

    # --- Report row ---
    $reportRows += [PSCustomObject]@{
        FirstName   = $firstName
        LastName    = $lastName
        Email       = $displayEmail
        Roles       = $rolesSummary
        Divisions   = $divsSummary
        Teams       = $teamsSummary
        SafeHaven   = if ($checkSafeHaven)  { if ($doneSafeHaven)  {'Y'} else {'N'} } else { 'N/A' }
        Concussion  = if ($checkConcussion) { if ($doneConcussion) {'Y'} else {'N'} } else { 'N/A' }
        SCA         = if ($checkSCA)        { if ($doneSCA)        {'Y'} else {'N'} } else { 'N/A' }
        SafeSport   = if ($checkSafeSport)  { if ($doneSafeSport)  {'Y'} else {'N'} } elseif (-not $isAdult) { 'Under 18' } else { 'N/A' }
        RiskStatus  = $riskStatus
        CoachTraining = if ($isCoachRole) {
                            if ($null -eq $coachTrainingStatus)     { 'Unknown Div' }
                            elseif ($coachTrainingStatus -eq $true) { "Current ($coachCurrentLevel)" }
                            else                                     { "Needed ($coachRequiredLabel)" }
                        } else { 'N/A' }
        RefTraining   = if ($isBlueStatus -or $isYouthRefRole) {
                            if ($needs8UOfficial) { "Needs 8U Official (Grade: none)" }
                            else                  { "8U Official OK (Grade: $refGrade)" }
                        } else { 'N/A' }
        LowerDivVerify = if ($needsCoach6U -and $needsCoach8U) { 'Needs 6U + 8U Coach course' }
                         elseif ($needsCoach6U)                { 'Needs 6U Coach course' }
                         elseif ($needsCoach8U)                { 'Needs 8U Coach course' }
                         else                                   { 'N/A' }
        Missing     = if ($anyIncomplete) { $missingList -join ', ' } else { '' }
        EmailSent   = 'Pending'
    }

    if (-not $anyIncomplete) {
        $phone = ($matchingRecords | Select-Object -First 1).Phone
        $divs  = ($matchingRecords | Where-Object { $_.Division -and $_.Division -ne 'Unallocated' } | Select-Object -ExpandProperty Division | Sort-Object -Unique) -join '; '
        $teams = ($matchingRecords | Where-Object { $_.Team     -and $_.Team     -ne 'Unallocated' } | Select-Object -ExpandProperty Team     | Sort-Object -Unique) -join '; '
        $compliantList += [PSCustomObject]@{
            FirstName = $firstName; LastName = $lastName
            Roles = $rolesSummary; Divisions = $divs; Teams = $teams
            Email = $displayEmail; Phone = $phone
        }
        Write-Host "✅ $firstName $lastName ($rolesSummary) - compliant" -ForegroundColor Green
        $reportRows[-1].EmailSent = 'No - Compliant'

        if ($SendCompliantEmail -and $safetyComplete) {
            # ── Build safety-completion congratulatory email ─────────────────
            $safetyStatusRows = ""
            if ($checkRisk)       { $safetyStatusRows += "<tr><td>Background Check (Risk Status)</td><td>$(Get-Badge $true)</td></tr>`n" }
            if ($checkSafeHaven)  { $safetyStatusRows += "<tr><td>AYSO's Safe Haven</td><td>$(Get-Badge $true)</td></tr>`n" }
            if ($checkConcussion) { $safetyStatusRows += "<tr><td>CDC Concussion Awareness</td><td>$(Get-Badge $true)</td></tr>`n" }
            if ($checkSCA)        { $safetyStatusRows += "<tr><td>Sudden Cardiac Arrest</td><td>$(Get-Badge $true)</td></tr>`n" }
            if ($checkSafeSport)  { $safetyStatusRows += "<tr><td>SafeSport</td><td>$(Get-Badge $true)</td></tr>`n" }

            $compliantHtml = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { margin:0; padding:0; font-family:Arial,sans-serif; font-size:14px; color:#333; background:#f0f0f0; }
    .wrapper { max-width:640px; margin:20px auto; background:#fff; border-radius:8px;
               overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.12); }
    .header  { background:#1a7a1a; padding:22px 28px; color:#fff; }
    .header h1 { margin:0; font-size:20px; font-weight:bold; }
    .header p  { margin:4px 0 0 0; font-size:13px; opacity:0.85; }
    .body    { padding:24px 28px; }
    .role-box { background:#f0f4ff; border:1px solid #c8d4f0; border-radius:6px;
                padding:10px 16px; margin:0 0 20px 0; font-size:13px; }
    .role-box ul { margin:4px 0 0 0; padding-left:18px; }
    .role-box li { margin-bottom:2px; }
    .status-table { width:100%; border-collapse:collapse; margin:16px 0 24px 0; }
    .status-table th { background:#1a7a1a; color:#fff; text-align:left;
                       padding:8px 12px; font-size:13px; }
    .status-table td { padding:8px 12px; border-bottom:1px solid #e0e0e0; font-size:13px; }
    .status-table tr:last-child td { border-bottom:none; }
    .status-table tr:nth-child(even) td { background:#f8f8f8; }
    a { color:#003366; }
    .footer { padding:14px 28px; background:#f0f0f0; font-size:12px;
              color:#666; border-top:1px solid #ddd; }
  </style>
</head>
<body>
<div class="wrapper">
  <div class="header">
    <h1>AYSO Region 138 - Safety Training Complete!</h1>
  </div>
  <div class="body">
    <p>Hi $firstName,</p>
    <p>Great news &#127881; &mdash; you have completed all required safety training for AYSO Region 138.
       Thank you for taking the time to get this done. Here is your current safety training status:</p>

    <table class="status-table">
      <thead>
        <tr><th>Safety Requirement</th><th>Status</th></tr>
      </thead>
      <tbody>
$safetyStatusRows      </tbody>
    </table>

    <p>You&rsquo;re all set on safety training. If you have any questions or believe something looks
       incorrect, just reply to this email and we&rsquo;ll sort it out.</p>

    <p>Thanks again for everything you do for the kids in our community!</p>

    <p>
      <strong>John Rennemeyer</strong><br>
      Coach Administrator<br>
      <a href="https://www.ayso138.org">AYSO Region 138</a>
    </p>
  </div>
  <div class="footer">
    Sent to $displayEmail on behalf of AYSO Region 138, Brigham City, Utah.<br>
    If you received this in error, please reply and let us know.
  </div>
</div>
</body>
</html>
"@
            $safeName2  = "$($lastName)_$($firstName)_SafetyComplete" -replace '[^a-zA-Z0-9_]', '_'
            $htmlPath2  = Join-Path $OutputFolder "$safeName2.html"
            $compliantHtml | Out-File -FilePath $htmlPath2 -Encoding utf8

            $toAddress2 = if ($TestMode) { $GmailFrom } else { $displayEmail }
            Write-Host "✉️  COMPLIANT | $firstName $lastName [$rolesSummary]$(if ($TestMode) { " → $GmailFrom" })" -ForegroundColor Green
            Write-Host "             Preview: $htmlPath2"

            if (-not $DryRun) {
                $mailParams2 = @{
                    From       = "AYSO 138 Coach Admin <$GmailFrom>"
                    To         = $toAddress2
                    Subject    = "You're All Set: AYSO Region 138 Safety Training Complete"
                    Body       = $compliantHtml
                    BodyAsHtml = $true
                    SmtpServer = "smtp.gmail.com"
                    Port       = 587
                    UseSsl     = $true
                    Credential = New-Object System.Management.Automation.PSCredential(
                                    $GmailFrom,
                                    (ConvertTo-SecureString $GmailPassword -AsPlainText -Force))
                }
                if ($CC) { $mailParams2.Cc = $CC }
                try {
                    Send-MailMessage @mailParams2
                    $sentCount++
                    $reportRows[-1].EmailSent = if ($TestMode) { "Compliant Email Sent (Test)" } else { "Compliant Email Sent" }
                    $logDir = Split-Path $SentLogPath -Parent
                    if ($logDir -and -not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
                    Add-Content -Path $SentLogPath -Value $displayEmail.ToLower()
                } catch {
                    $errorCount++
                    $reportRows[-1].EmailSent = "Compliant Email FAILED"
                    Write-Host "  ❌ FAILED: $($_.Exception.Message)" -ForegroundColor Red
                }
                if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
            } else {
                $sentCount++
                $reportRows[-1].EmailSent = "Dry Run (Compliant)"
            }
        } else {
            $skippedCount++
        }
        continue
    }

    # ── Safety-complete but still has non-safety items outstanding ────────────────
    # Send compliant email if requested, then fall through to reminder email only
    # if there are non-safety items still outstanding (e.g. coach training).
    if ($SendCompliantEmail -and $safetyComplete) {
        $phone = ($matchingRecords | Select-Object -First 1).Phone
        $divs  = ($matchingRecords | Where-Object { $_.Division -and $_.Division -ne 'Unallocated' } | Select-Object -ExpandProperty Division | Sort-Object -Unique) -join '; '
        $teams = ($matchingRecords | Where-Object { $_.Team     -and $_.Team     -ne 'Unallocated' } | Select-Object -ExpandProperty Team     | Sort-Object -Unique) -join '; '
        $compliantList += [PSCustomObject]@{
            FirstName = $firstName; LastName = $lastName
            Roles = $rolesSummary; Divisions = $divs; Teams = $teams
            Email = $displayEmail; Phone = $phone
        }
        $safetyStatusRows = ""
        if ($checkRisk)       { $safetyStatusRows += "<tr><td>Background Check (Risk Status)</td><td>$(Get-Badge $true)</td></tr>`n" }
        if ($checkSafeHaven)  { $safetyStatusRows += "<tr><td>AYSO's Safe Haven</td><td>$(Get-Badge $true)</td></tr>`n" }
        if ($checkConcussion) { $safetyStatusRows += "<tr><td>CDC Concussion Awareness</td><td>$(Get-Badge $true)</td></tr>`n" }
        if ($checkSCA)        { $safetyStatusRows += "<tr><td>Sudden Cardiac Arrest</td><td>$(Get-Badge $true)</td></tr>`n" }
        if ($checkSafeSport)  { $safetyStatusRows += "<tr><td>SafeSport</td><td>$(Get-Badge $true)</td></tr>`n" }

        $compliantHtml = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { margin:0; padding:0; font-family:Arial,sans-serif; font-size:14px; color:#333; background:#f0f0f0; }
    .wrapper { max-width:640px; margin:20px auto; background:#fff; border-radius:8px;
               overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.12); }
    .header  { background:#1a7a1a; padding:22px 28px; color:#fff; }
    .header h1 { margin:0; font-size:20px; font-weight:bold; }
    .header p  { margin:4px 0 0 0; font-size:13px; opacity:0.85; }
    .body    { padding:24px 28px; }
    .status-table { width:100%; border-collapse:collapse; margin:16px 0 24px 0; }
    .status-table th { background:#1a7a1a; color:#fff; text-align:left;
                       padding:8px 12px; font-size:13px; }
    .status-table td { padding:8px 12px; border-bottom:1px solid #e0e0e0; font-size:13px; }
    .status-table tr:last-child td { border-bottom:none; }
    .status-table tr:nth-child(even) td { background:#f8f8f8; }
    a { color:#003366; }
    .footer { padding:14px 28px; background:#f0f0f0; font-size:12px;
              color:#666; border-top:1px solid #ddd; }
  </style>
</head>
<body>
<div class="wrapper">
  <div class="header">
    <h1>AYSO Region 138 - Safety Training Complete!</h1>
  </div>
  <div class="body">
    <p>Hi $firstName,</p>
    <p>Great news &#127881; &mdash; you have completed all required safety training for AYSO Region 138.
       Thank you for taking the time to get this done. Here is your current safety training status:</p>

    <table class="status-table">
      <thead>
        <tr><th>Safety Requirement</th><th>Status</th></tr>
      </thead>
      <tbody>
$safetyStatusRows      </tbody>
    </table>

    <p>You&rsquo;re all set on safety training. If you have any questions or believe something looks
       incorrect, just reply to this email and we&rsquo;ll sort it out.</p>

    <p>Thanks again for everything you do for the kids in our community!</p>

    <p>
      <strong>John Rennemeyer</strong><br>
      Coach Administrator<br>
      <a href="https://www.ayso138.org">AYSO Region 138</a>
    </p>
  </div>
  <div class="footer">
    Sent to $displayEmail on behalf of AYSO Region 138, Brigham City, Utah.<br>
    If you received this in error, please reply and let us know.
  </div>
</div>
</body>
</html>
"@
        $safeName2 = "$($lastName)_$($firstName)_SafetyComplete" -replace '[^a-zA-Z0-9_]', '_'
        $htmlPath2 = Join-Path $OutputFolder "$safeName2.html"
        $compliantHtml | Out-File -FilePath $htmlPath2 -Encoding utf8

        $toAddress2 = if ($TestMode) { $GmailFrom } else { $displayEmail }
        Write-Host "✉️  COMPLIANT | $firstName $lastName [$rolesSummary]$(if ($TestMode) { " → $GmailFrom" })" -ForegroundColor Green
        Write-Host "             Preview: $htmlPath2"

        if (-not $DryRun) {
            $mailParams2 = @{
                From       = "AYSO 138 Coach Admin <$GmailFrom>"
                To         = $toAddress2
                Subject    = "You're All Set: AYSO Region 138 Safety Training Complete"
                Body       = $compliantHtml
                BodyAsHtml = $true
                SmtpServer = "smtp.gmail.com"
                Port       = 587
                UseSsl     = $true
                Credential = New-Object System.Management.Automation.PSCredential(
                                $GmailFrom,
                                (ConvertTo-SecureString $GmailPassword -AsPlainText -Force))
            }
            if ($CC) { $mailParams2.Cc = $CC }
            try {
                Send-MailMessage @mailParams2
                $sentCount++
                $reportRows[-1].EmailSent = if ($TestMode) { "Compliant Email Sent (Test)" } else { "Compliant Email Sent" }
                $logDir = Split-Path $SentLogPath -Parent
                if ($logDir -and -not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
                Add-Content -Path $SentLogPath -Value $displayEmail.ToLower()
            } catch {
                $errorCount++
                $reportRows[-1].EmailSent = "Compliant Email FAILED"
                Write-Host "  ❌ FAILED: $($_.Exception.Message)" -ForegroundColor Red
            }
            if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
        } else {
            $sentCount++
            $reportRows[-1].EmailSent = "Dry Run (Compliant)"
        }
        # Do NOT continue — fall through to the reminder email for remaining non-safety items
    }

    # ========================================================================
    # BUILD EMAIL
    # ========================================================================

    # --- Role/team context block ---
    $roleLines = ""
    foreach ($rec in ($matchingRecords | Sort-Object Role)) {
        $teamDisplay = if ($rec.Team -ne 'Unallocated') { " - $($rec.Team)" } else { "" }
        $roleLines += "<li><strong>$($rec.Role)</strong>$teamDisplay <span style='color:#666;font-size:12px;'>($($rec.Division))</span></li>`n"
    }

    # --- Status table rows ---
    $statusRows = ""
    if ($checkRisk)       { $statusRows += "<tr><td>Background Check (Risk Status)</td><td>$(Get-Badge $doneRisk)</td></tr>`n" }
    if ($checkSafeHaven)  { $statusRows += "<tr><td>AYSO's Safe Haven</td><td>$(Get-Badge $doneSafeHaven)</td></tr>`n" }
    if ($checkConcussion) { $statusRows += "<tr><td>CDC Concussion Awareness</td><td>$(Get-Badge $doneConcussion)</td></tr>`n" }
    if ($checkSCA)        { $statusRows += "<tr><td>Sudden Cardiac Arrest</td><td>$(Get-Badge $doneSCA)</td></tr>`n" }
    if ($checkSafeSport)  { $statusRows += "<tr><td>SafeSport</td><td>$(Get-Badge $doneSafeSport)</td></tr>`n" }
    
    # Also show N/A rows for adult-only items if the volunteer is under 18
    if (($requiredItems -contains 'SafeSport') -and (-not $isAdult)) {
        $statusRows += "<tr><td>SafeSport</td><td><span style='color:#888;'>N/A - not required under 18</span></td></tr>`n"
    }
    if (($requiredItems -contains 'Risk') -and (-not $isAdult)) {
        $statusRows += "<tr><td>Background Check</td><td><span style='color:#888;'>N/A - not required under 18</span></td></tr>`n"
    }

    # Coach training row - Head Coach / Assistant Coach only (suppressed in SafetyOnly mode)
    if ($isCoachRole -and -not $SafetyOnly) {
        $ctCell = if ($null -eq $coachTrainingStatus) {
            "<span style='color:#888888;'>N/A - division unrecognized</span>"
        } elseif ($coachTrainingStatus -eq $true) {
            "<span style='color:#1a7a1a;font-weight:bold;'>&#10003; Current - $coachCurrentLevel</span>"
        } else {
            $displayLevel = if ($coachCurrentLevel -ne 'None') { " <br>&#10003; Currently have $coachCurrentLevel license" } else { '' }
            "<span style='color:#cc0000;font-weight:bold;'>&#10007; Needed - $coachRequiredLabel$displayLevel</span>"
        }
        $statusRows += "<tr><td>Coach Training</td><td>$ctCell</td></tr>`n"
    }

    # Referee training rows - Youth Referees only (suppressed in SafetyOnly mode)
    if ((-not $SafetyOnly) -and ($isBlueStatus -or $isYouthRefRole)) {
        $refCell = if ($needs8UOfficial) {
            "<span style='color:#cc0000;font-weight:bold;'>&#10007; Needed</span>"
        } else {
            "<span style='color:#1a7a1a;font-weight:bold;'>&#10003; Current - $refGrade</span>"
        }
        $statusRows += "<tr><td>Referee Training (8U Official)</td><td>$refCell</td></tr>`n"
    }

    # (Lower-division coach course gaps appear in the instruction section only, no status row)

    # --- Instruction sections (only for missing items) ---
    $instructions = ""

    if ($needsRisk) {
        $currentStatus = if ($riskStatus -and $riskStatus -ne 'None') { $riskStatus } else { "not on file" }
        $instructions += @"
    <div class="section">
      <h3>&#128196; Background Check - Steps to Complete</h3>
      <p>You currently do not have a completed background check. You <strong>must have a completed background check</strong> before the season begins.</p>
      <ol>
        <li>Log in to <a href="https://www.ayso138.org">www.ayso138.org</a>.</li>
        <li>Click <strong>My Account</strong> (top right), then select the <strong>Volunteer</strong> tab on the left.</li>
        <li>If Risk Status shows empty or expired, check the box and click <strong>Renew &amp; Update</strong>.<br>
            <em>Do not start a new background check if the expiration date has not passed.</em></li>
        <li>Confirm your information using your full legal name and submit.</li>
        <li>Within an hour, you should receive an email from <strong>noreply@sterlingcheck.com</strong> with
            the subject &ldquo;Your AYSO Background Check is Incomplete - Please submit now.&rdquo;</li>
        <li>Click the link in that email and follow the instructions to completion.</li>
        <li>If you don&rsquo;t receive the Sterling email, check your spam/junk folder.<br>
            For help: call <strong>855-326-1860 (Option 3)</strong> or email
            <a href="mailto:TheAdvocates@SterlingVolunteers.com">TheAdvocates@SterlingVolunteers.com</a>.</li>
      </ol>
    </div>
"@
    }

    $needsAnyTraining = $needsSafeHaven -or $needsConcussion -or $needsSCA -or $needsSafeSport
    if ($needsAnyTraining) {
        $courseList = ""
        if ($needsSafeHaven)  { $courseList += "        <li>AYSO&rsquo;s Safe Haven</li>`n" }
        if ($needsConcussion) { $courseList += "        <li>CDC Concussion Awareness</li>`n" }
        if ($needsSCA)        { $courseList += "        <li>Sudden Cardiac Arrest</li>`n" }
        if ($needsSafeSport)  { $courseList += "        <li>SafeSport</li>`n" }

        $instructions += @"
    <div class="section">
      <h3>&#127891; Online Training - Steps to Complete</h3>
      <ol>
        <li>Log into your Sports Connect account at <a href="https://www.ayso138.org">www.ayso138.org</a>.</li>
        <li>Click <strong>My Account</strong>, if needed.</li>
        <li>Click <strong>AYSOU</strong> in the left menu.</li>
        <li>Once <strong>AYSOU loads</strong>, click <strong>Training Library</strong> in the left menu.</li>
        <li>Click <strong>View Courses</strong> on the <strong>Safe Haven</strong> card and complete the following missing course(s):
          <ul>
$courseList          </ul>
        </li>
      </ol>
    </div>
"@
    }

    # ── Referee Training Section (Youth Referees only) ────────────────────────
    if ($needsRefTraining -and $needs8UOfficial) {
        $instructions += @"
    <div class="section">
      <h3>&#9917; Referee Training - Steps to Complete</h3>
      <p>Youth referees need to complete the <strong>8U Official - Full Online Course</strong> in AYSOU under the <strong>Referee</strong> card:</p>
      <ol>
        <li>Log into your Sports Connect account at <a href="https://www.ayso138.org">www.ayso138.org</a>.</li>
        <li>Click <strong>My Account</strong>, if needed.</li>
        <li>Click <strong>AYSOU</strong> in the left menu.</li>
        <li>Once <strong>AYSOU</strong> loads, click <strong>Training Library</strong>.</li>
        <li>Open the <strong>Referee</strong> card and enroll in <strong>8U Official - Full Online Course</strong>.</li>
        <li>Complete the course.</li>
      </ol>      
    </div>
"@
    }

    # ── Lower-Division Coach Course Section ────────────────────────────────────
    if ($needsCoach6U) {
        $instructions += @"
    <div class="section">
      <h3>&#127941; 6U Coach Course - Action Required</h3>
      <p>You are coaching a 6U division but do not have the <strong>6U Coach - Full Online Course</strong> on file. Please complete it in AYSOU.</p>
      <ol>
        <li>Log into your Sports Connect account at <a href="https://www.ayso138.org">www.ayso138.org</a>.</li>
        <li>Click <strong>My Account</strong>, if needed.</li>
        <li>Click <strong>AYSOU</strong> in the left menu.</li>
        <li>Once AYSOU loads, click <strong>Training Library</strong>.</li>
        <li>Open the <strong>Coaches</strong> card and enroll in <strong>6U Coach - Full Online Course</strong>.</li>
        <li>Complete the course.</li>
      </ol>
    </div>
"@
    }

#     if ($needsCoach8U) {
#         $instructions += @"
#     <div class="section">
#       <h3>&#127941; 8U Coach Course - Action Required</h3>
#       <p>You are coaching an 8U division but do not have the <strong>8U Coach - Full Online Course</strong> on file. Please complete it in AYSOU.</p>
#       <ol>
#         <li>Log into your Sports Connect account at <a href="https://www.ayso138.org">www.ayso138.org</a>.</li>
#         <li>Click <strong>My Account</strong>, if needed.</li>
#         <li>Click <strong>AYSOU</strong> in the left menu.</li>
#         <li>Once AYSOU loads, click <strong>Training Library</strong>.</li>
#         <li>Open the <strong>Coaches</strong> card and enroll in <strong>8U Coach - Full Online Course</strong>.</li>
#         <li>Complete the course.</li>
#       </ol>
#     </div>
# "@
#     }

    # ── Coach Training Section (Head Coach / Assistant Coach only) ────────────
    if ($isCoachRole -and $needsCoachTraining) {
        if ($coachIsJHHS) {
            # Build level-specific intro and course steps based on where the coach currently is.
            # Target: Intermediate (14U) Coach. 16U-19U Advanced requires full completion
            # of the 14U in-person course first - do not reference it here.
            $jhIntro   = ''
            $jhCourses = ''

            if ($coachCurrentRank -eq 0) {
                $jhIntro = "You do not currently have a coaching license on file. For Junior High/High School coaching, the path runs through two certification levels: <strong>12U Coach</strong> and then <strong>Intermediate (14U) Coach</strong>. Both online portions can be completed now - no in-person event is required to get started."
                $jhCourses = @"
        <li>Open the <strong>Coaches</strong> card.</li>
        <li>Enroll in <strong>12U Coach - Online + In-Person Course</strong> and complete the <em>Online portion</em>.</li>
        <li>Once the 12U online portion is done, return to the Training Library and enroll in <strong>Intermediate (14U) Coach - Online + In-Person Course</strong> and complete the <em>Online portion</em>.</li>
"@
            } elseif ($coachCurrentRank -le 3) {
                $jhIntro = "Your current license (<strong>$coachCurrentLevel</strong>) is a great start. For Junior High/High School coaching, the next steps are the <strong>12U Coach</strong> course and then the <strong>Intermediate (14U) Coach</strong> course. Both online portions can be completed now - no in-person event needed."
                $jhCourses = @"
        <li>Open the <strong>Coaches</strong> card.</li>
        <li>Enroll in <strong>12U Coach - Online + In-Person Course</strong> and complete the <em>Online portion</em>.</li>
        <li>Once the 12U online portion is done, return to the Training Library and enroll in <strong>Intermediate (14U) Coach - Online + In-Person Course</strong> and complete the <em>Online portion</em>.</li>
"@
            } elseif ($coachCurrentRank -eq 4) {
                $jhIntro = "You have your <strong>10U Coach</strong> license - great progress! For Junior High/High School coaching, the next steps are the <strong>12U Coach</strong> course and then the <strong>Intermediate (14U) Coach</strong> course. Both online portions can be completed now - no in-person event needed."
                $jhCourses = @"
        <li>Open the <strong>Coaches</strong> card.</li>
        <li>Enroll in <strong>12U Coach - Online + In-Person Course</strong> and complete the <em>Online portion</em>.</li>
        <li>Once the 12U online portion is done, return to the Training Library and enroll in <strong>Intermediate (14U) Coach - Online + In-Person Course</strong> and complete the <em>Online portion</em>.</li>
"@
            } elseif ($coachCurrentRank -eq 5) {
                $jhIntro = "You have your <strong>12U Coach</strong> certification - great foundation! The next step for Junior High/High School coaching is the <strong>Intermediate (14U) Coach</strong> course. You can complete the online portion now without waiting for an in-person event."
                $jhCourses = @"
        <li>Open the <strong>Coaches</strong> card.</li>
        <li>Enroll in <strong>Intermediate (14U) Coach - Online + In-Person Course</strong> and complete the <em>Online portion</em>.</li>
"@
            }

            $instructions += @"
    <div class="section">
      <h3>&#127941; Coach Training - Junior High/High School</h3>
      <p>$jhIntro</p>
      <p>Please complete the online portion now, so you'll be better prepared for your players right away. The in-person portion is still required, and you can attend a session when one becomes available that fits your schedule.</p>
      <p>To enroll:</p>
      <ol>
        <li>Log into your Sports Connect account at <a href="https://www.ayso138.org">www.ayso138.org</a>.</li>
        <li>Click <strong>My Account</strong>, if needed.</li>
        <li>Click <strong>AYSOU</strong> in the left menu.</li>
        <li>Once AYSOU loads, click <strong>Training Library</strong>.</li>
$jhCourses      </ol>
    </div>
"@
        } else {
            # ── 10U: target is 10U Coach (full online) ───────────────────────────────
            if ($coachHighestAge -eq 10) {
                $divIntro = ''

                if ($coachCurrentRank -eq 0) {
                    $divIntro = "You do not currently have a coaching license on file. For 10U coaching, you need the <strong>$($coachCourseLabel)</strong>."
                } elseif ($coachCurrentRank -eq 1) {
                    $divIntro = "You have your <strong>Kickstart Soccer Play Leader</strong> license; great start! For 10U coaching, the next step is the <strong>$($coachCourseLabel)</strong>."
                } elseif ($coachCurrentRank -eq 2) {
                    $divIntro = "You have your <strong>6U Coach</strong> license; great start! For 10U coaching, the next step is the <strong>$($coachCourseLabel)</strong>."
                } elseif ($coachCurrentRank -eq 3) {
                    $divIntro = "You have your <strong>8U Coach</strong> license; you&rsquo;re one step away! For 10U coaching, the next step is the <strong>$($coachCourseLabel)</strong>."
                }

                $instructions += @"
    <div class="section">
      <h3>&#127941; Coach Training - Action Required</h3>
      <p>$divIntro</p>
      <p>Please complete the online portion now, so you'll be better prepared for your players right away. The in-person portion is still required, and you can attend a session when one becomes available that fits your schedule.</p>
      <p>To enroll:</p>
      <ol>
        <li>Log into your Sports Connect account at <a href="https://www.ayso138.org">www.ayso138.org</a>.</li>
        <li>Click <strong>My Account</strong>, if needed.</li>
        <li>Click <strong>AYSOU</strong> in the left menu.</li>
        <li>Once AYSOU loads, click <strong>Training Library</strong>.</li>
        <li>Open the <strong>Coaches</strong> card.</li>
        <li>Click <strong>Enroll</strong> under <strong>$($coachCourseLabel)</strong> and complete the <em>Online portion</em>.</li>
      </ol>
    </div>
"@

            # ── 12U: target is 12U Coach (rank 5), online + in-person ────────────────
            } elseif ($coachHighestAge -eq 12) {
                $divIntro = ''

                if ($coachCurrentRank -eq 0) {
                    $divIntro = "You do not currently have a coaching license on file. For 12U coaching, you need the <strong>$($coachCourseLabel)</strong>. Start with the online portion now."
                } elseif ($coachCurrentRank -eq 1) {
                    $divIntro = "You have your <strong>Kickstart Soccer Play Leader</strong> license; great start! For 12U coaching, the next step is the <strong>$($coachCourseLabel)</strong>."
                } elseif ($coachCurrentRank -eq 2) {
                    $divIntro = "You have your <strong>6U Coach</strong> license; great start! For 12U coaching, the next step is the <strong>$($coachCourseLabel)</strong>."
                } elseif ($coachCurrentRank -eq 3) {
                    $divIntro = "You have your <strong>8U Coach</strong> license; solid foundation! For 12U coaching, the next step is the <strong>$($coachCourseLabel)</strong>."
                } elseif ($coachCurrentRank -eq 4) {
                    $divIntro = "You have your <strong>10U Coach</strong> license; great progress! The next and final step for 12U coaching is the <strong>$($coachCourseLabel)</strong>."
                }

                $instructions += @"
    <div class="section">
      <h3>&#127941; Coach Training - Action Required</h3>
      <p>$divIntro</p>
      <p>Please complete the online portion now, so you'll be better prepared for your players right away. The in-person portion is still required, and you can attend a session when one becomes available that fits your schedule.</p>
      <p>To enroll:</p>
      <ol>
        <li>Log into your Sports Connect account at <a href="https://www.ayso138.org">www.ayso138.org</a>.</li>
        <li>Click <strong>My Account</strong>, if needed.</li>
        <li>Click <strong>AYSOU</strong> in the left menu.</li>
        <li>Once AYSOU loads, click <strong>Training Library</strong>.</li>
        <li>Open the <strong>Coaches</strong> card.</li>
        <li>Enroll in <strong>$($coachCourseLabel)</strong>.</li>
        <li><strong>Complete the online portion</strong> of the course.</li>
      </ol>
    </div>
"@

            # ── Playground / Schoolyard / Kickstart: Kickstart Soccer Play Leader ──
            } elseif ($coachHighestAge -eq 0) {
                $instructions += @"
    <div class="section">
      <h3>&#127941; Coach Training - Action Required</h3>
      <p>For Playground, Schoolyard, and Kickstart coaching, you need to complete the <strong>Kickstart Soccer Play Leader - Full Online Course</strong>.</p>
      <p>To enroll:</p>
      <ol>
        <li>Log into your Sports Connect account at <a href="https://www.ayso138.org">www.ayso138.org</a>.</li>
        <li>Click <strong>My Account</strong>, if needed.</li>
        <li>Click <strong>AYSOU</strong> in the left menu.</li>
        <li>Once AYSOU loads, click <strong>Training Library</strong>.</li>
        <li>Open the <strong>Coaches</strong> card.</li>
        <li>Click <strong>Enroll</strong> under <strong>Kickstart Soccer Play Leader - Full Online Course</strong>.</li>
        <li><strong>Complete</strong> the course.</li>
      </ol>
    </div>
"@

            # ── 6U / 8U: single course, no stepped path needed ────────────────────────
            } else {
                $levelNote = if ($coachCurrentLevel -ne 'None') {
                    "Your current coaching license is <strong>$coachCurrentLevel</strong>, which does not yet cover your highest coaching division."
                } else {
                    "You do not currently have a coaching license on file for AYSO Region 138."
                }

                $instructions += @"
    <div class="section">
      <h3>&#127941; Coach Training - Action Required</h3>
      <p>$levelNote Your coaching assignment requires the <strong>$coachCourseLabel</strong>.</p>
      <p>To enroll in your coach training course:</p>
      <ol>
        <li>Log into your Sports Connect account at <a href="https://www.ayso138.org">www.ayso138.org</a>.</li>
        <li>Click <strong>My Account</strong>, if needed.</li>
        <li>Click <strong>AYSOU</strong> in the left menu.</li>
        <li>Once <strong>AYSOU</strong> loads, click <strong>Training Library</strong> in the left menu.</li>
        <li>Click <strong>View Courses</strong> on the <strong>Coaches</strong> card.</li>
        <li>Click <strong>Enroll</strong> under <strong>$coachCourseLabel</strong>.</li>
        <li><strong>Complete</strong> the course.</li>
      </ol>
    </div>
"@
            }
        }
    }

    # --- Assemble full HTML ---
    $htmlBody = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { margin:0; padding:0; font-family:Arial,sans-serif; font-size:14px; color:#333; background:#f0f0f0; }
    .wrapper { max-width:640px; margin:20px auto; background:#fff; border-radius:8px;
               overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.12); }
    .header  { background:#003366; padding:22px 28px; color:#fff; }
    .header h1 { margin:0; font-size:20px; font-weight:bold; }
    .header p  { margin:4px 0 0 0; font-size:13px; opacity:0.85; }
    .body    { padding:24px 28px; }
    .role-box { background:#f0f4ff; border:1px solid #c8d4f0; border-radius:6px;
                padding:10px 16px; margin:0 0 20px 0; font-size:13px; }
    .role-box ul { margin:4px 0 0 0; padding-left:18px; }
    .role-box li { margin-bottom:2px; }
    .status-table { width:100%; border-collapse:collapse; margin:16px 0 24px 0; }
    .status-table th { background:#003366; color:#fff; text-align:left;
                       padding:8px 12px; font-size:13px; }
    .status-table td { padding:8px 12px; border-bottom:1px solid #e0e0e0; font-size:13px; }
    .status-table tr:last-child td { border-bottom:none; }
    .status-table tr:nth-child(even) td { background:#f8f8f8; }
    .section { background:#f5f8fc; border-left:4px solid #003366; border-radius:4px;
               padding:16px 20px; margin-bottom:20px; }
    .section h3 { margin:0 0 12px 0; color:#003366; font-size:15px; }
    .section ol { margin:0; padding-left:20px; }
    .section ol li { margin-bottom:8px; line-height:1.5; }
    .section ul { margin:8px 0 0 0; padding-left:20px; }
    .section ul li { margin-bottom:4px; }
    a { color:#003366; }
    .footer { padding:14px 28px; background:#f0f0f0; font-size:12px;
              color:#666; border-top:1px solid #ddd; }
  </style>
</head>
<body>
<div class="wrapper">
  <div class="header">
    <h1>AYSO Region 138 - Volunteer Compliance Reminder</h1>
  </div>
  <div class="body">
    <p>Hi $firstName,</p>
    <p>Thank you again for volunteering with AYSO Region 138! As a reminder, all volunteers need to complete
       the online portion of their required training as soon as possible.</p>
"@ + $(if ($noCredRecord) { @"
       <div class="section">
         <p><strong>Note:</strong> We could not find a credentials record for your account in our system. 
         This means your status below may not be accurate. Please complete all required training listed below, 
         and if you believe you have already completed some of these items, please reply to this email so we 
         can investigate and update your record.</p>
       </div>
"@ } elseif ($matchedByName) { @"
       <div class="section">
         <p><strong>Note:</strong> Your email address in our volunteer roster ($em) does not match 
         the email in our credentials system ($displayEmail). We matched you by name and are using your 
         credentials from $displayEmail. If this email address is incorrect, please reply to let us know.</p>
       </div>
"@ } else { "" }) + @"
       <p>Here is your current status:</p>

    <table class="status-table">
      <thead>
        <tr><th>Requirement</th><th>Status</th></tr>
      </thead>
      <tbody>
$statusRows      </tbody>
    </table>

    <p><strong>Please complete the item(s) marked
       <strong style="color:#cc0000;">&#10007; Needed</strong> following the instructions below:</strong></p>

$instructions

    <p>If you have completed any of the training recently, please allow 24&ndash;48 hours
       for your account at <a href="https://www.ayso138.org">www.ayso138.org</a>
       to reflect the training completion.</p>

    <p>Questions or issues? Reply to this email and we&rsquo;ll get you sorted out.
       We appreciate your time and dedication to the kids in our community!</p>

    <p>
      <strong>John Rennemeyer</strong><br>
      Coach Administrator<br>
      <a href="https://www.ayso138.org">AYSO Region 138</a>
    </p>
  </div>
  <div class="footer">
    Sent to $displayEmail on behalf of AYSO Region 138, Brigham City, Utah.<br>
    If you received this in error, please reply and let us know.
  </div>
</div>
</body>
</html>
"@

    # --- Build division/role suffix for filename ---
    $labelSortOrder = @{ 'PG'=0; 'SY'=1; 'KS'=2; '6U'=3; '8U'=4; '10U'=5; '12U'=6; 'JHHS'=7;
                         'HC'=8; 'AC'=9; 'YR'=10; 'Ref'=11; 'TM'=12; 'FS'=13; 'BM'=14 }
    $fileLabels = [System.Collections.Generic.List[string]]::new()

    # Coach divisions + role abbreviations
    if ($isCoachRole -and $allCoachRecs.Count -gt 0) {
        foreach ($rec in $allCoachRecs) {
            $lbl = Get-DivisionLabel $rec.Division
            if ($lbl -and $lbl -notin $fileLabels) { $fileLabels.Add($lbl) }
        }
        # Add HC / AC role abbreviations after the division labels
        $hasHC = @($matchingRecords | Where-Object { $_.Role -eq 'Head Coach' }).Count -gt 0
        $hasAC = @($matchingRecords | Where-Object { $_.Role -eq 'Assistant Coach' }).Count -gt 0
        if ($hasHC -and 'HC' -notin $fileLabels) { $fileLabels.Add('HC') }
        if ($hasAC -and 'AC' -notin $fileLabels) { $fileLabels.Add('AC') }
    }
    # Non-coach roles - add role abbreviation.
    # Blue risk status overrides the "Referee" role label → always "YR" for youth volunteers.
    foreach ($rec in $matchingRecords) {
        if ($rec.Role -notin @('Head Coach','Assistant Coach')) {
            $roleAbbr = switch ($rec.Role) {
                'Youth Referee' { 'YR'  }
                'Referee'       { if ($isBlueStatus) { 'YR' } else { 'Ref' } }
                'Team Manager'  { 'TM'  }
                'Field Setup'   { 'FS'  }
                'Board Member'  { 'BM'  }
                default         { ''    }
            }
            if ($roleAbbr -and $roleAbbr -notin $fileLabels) { $fileLabels.Add($roleAbbr) }
        }
    }
    $fileLabels = @($fileLabels | Sort-Object { if ($labelSortOrder.ContainsKey($_)) { $labelSortOrder[$_] } else { 99 } })
    $suffix   = if ($fileLabels.Count -gt 0) { '_' + ($fileLabels -join '_') } else { '' }
    $safeName = "$($lastName)_$($firstName)$suffix" -replace '[^a-zA-Z0-9_]', '_'
    $htmlPath = Join-Path $OutputFolder "$safeName.html"
    $htmlBody | Out-File -FilePath $htmlPath -Encoding utf8

    # --- Determine To / CC ---
    $toAddress = if ($TestMode) { $GmailFrom } else { $displayEmail }

    # Track for summary and CSV export
    $ncPhone = ($matchingRecords | Select-Object -First 1).Phone
    $ncDivs  = ($matchingRecords | Where-Object { $_.Division -and $_.Division -ne 'Unallocated' } | Select-Object -ExpandProperty Division | Sort-Object -Unique) -join '; '
    $ncTeams = ($matchingRecords | Where-Object { $_.Team     -and $_.Team     -ne 'Unallocated' } | Select-Object -ExpandProperty Team     | Sort-Object -Unique) -join '; '
    $nonCompliantList += [PSCustomObject]@{
        FirstName = $firstName; LastName = $lastName
        Roles = $rolesSummary; Divisions = $ncDivs; Teams = $ncTeams
        Email = $displayEmail; Phone = $ncPhone
        MissingTraining = ($missingList -join ', ')
    }

    # --- Console output ---
    $modeTag = if ($TestMode) { " → $GmailFrom" } else { "" }
    if ($DryRun) {
        Write-Host "📝 DRY RUN  | $firstName $lastName [$rolesSummary]" -ForegroundColor Yellow
    } else {
        Write-Host "📧 SENDING  | $firstName $lastName [$rolesSummary]$modeTag" -ForegroundColor Cyan
    }
    Write-Host "            Missing: $($missingList -join ', ')"
    Write-Host "            Preview: $htmlPath"

    if (-not $DryRun) {
        $mailParams = @{
            From       = "AYSO 138 Coach Admin <$GmailFrom>"
            To         = $toAddress
            Subject    = "Action Required: AYSO Region 138 Volunteer Compliance Needed"
            Body       = $htmlBody
            BodyAsHtml = $true
            SmtpServer = "smtp.gmail.com"
            Port       = 587
            UseSsl     = $true
            Credential = New-Object System.Management.Automation.PSCredential(
                            $GmailFrom,
                            (ConvertTo-SecureString $GmailPassword -AsPlainText -Force))
        }
        if ($CC) { $mailParams.Cc = $CC }

        try {
            Send-MailMessage @mailParams
            $sentCount++
            $reportRows[-1].EmailSent = if ($TestMode) { "Sent (Test)" } else { "Sent" }

            # Log the address so future runs with -SkipAlreadySent can skip it
            $logDir = Split-Path $SentLogPath -Parent
            if ($logDir -and -not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
            Add-Content -Path $SentLogPath -Value $displayEmail.ToLower()

        } catch {
            $errorCount++
            $reportRows[-1].EmailSent = "FAILED"
            Write-Host "  ❌ FAILED: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Throttle: pause between sends to avoid Gmail rate limiting
        if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
    } else {
        $reportRows[-1].EmailSent = "Dry Run"
        $sentCount++
    }

}  # end foreach volunteer

# ============================================================================
# EXPORT CSV REPORT
# ============================================================================

if ($ExportReport) {
    $reportPath = Join-Path $OutputFolder "ComplianceReport.csv"
    $reportRows | Export-Csv -Path $reportPath -NoTypeInformation -Encoding utf8
    Write-Host "`n📊 Report saved: $reportPath" -ForegroundColor Magenta
}

# Always write the compliant / non-compliant CSVs
$compliantCsvPath    = Join-Path $OutputFolder "Compliant.csv"
$nonCompliantCsvPath = Join-Path $OutputFolder "NonCompliant.csv"

if ($compliantList.Count -gt 0) {
    $compliantList |
        Select-Object FirstName, LastName, Roles, Divisions, Teams, Email, Phone |
        Sort-Object LastName, FirstName |
        Export-Csv -Path $compliantCsvPath -NoTypeInformation -Encoding utf8
} else {
    # Write an empty file with headers so callers always have a file to reference
    [PSCustomObject]@{ FirstName=''; LastName=''; Roles=''; Divisions=''; Teams=''; Email=''; Phone='' } |
        Export-Csv -Path $compliantCsvPath -NoTypeInformation -Encoding utf8
    # Remove the single blank data row, keep headers only
    (Get-Content $compliantCsvPath | Select-Object -First 1) | Set-Content $compliantCsvPath
}
Write-Host "📋 Compliant CSV    : $compliantCsvPath" -ForegroundColor Green

if ($nonCompliantList.Count -gt 0) {
    $nonCompliantList |
        Select-Object FirstName, LastName, Roles, Divisions, Teams, Email, Phone, MissingTraining |
        Sort-Object LastName, FirstName |
        Export-Csv -Path $nonCompliantCsvPath -NoTypeInformation -Encoding utf8
} else {
    [PSCustomObject]@{ FirstName=''; LastName=''; Roles=''; Divisions=''; Teams=''; Email=''; Phone=''; MissingTraining='' } |
        Export-Csv -Path $nonCompliantCsvPath -NoTypeInformation -Encoding utf8
    (Get-Content $nonCompliantCsvPath | Select-Object -First 1) | Set-Content $nonCompliantCsvPath
}
Write-Host "📋 Non-Compliant CSV: $nonCompliantCsvPath" -ForegroundColor Yellow

# ============================================================================
# FINAL SUMMARY
# ============================================================================

Write-Host ""
Write-Host "  Fully compliant : $($compliantList.Count)"      -ForegroundColor Green
if ($compliantList.Count -gt 0) {
    foreach ($cv in ($compliantList | Sort-Object LastName, FirstName)) {
        Write-Host "    ✅ $($cv.FirstName) $($cv.LastName)  [$($cv.Roles)]" -ForegroundColor Green
    }
}
Write-Host "  Non-compliant   : $($nonCompliantList.Count)" -ForegroundColor $(if ($nonCompliantList.Count -gt 0) {"Yellow"} else {"Gray"})
if ($nonCompliantList.Count -gt 0) {
    foreach ($nc in ($nonCompliantList | Sort-Object LastName, FirstName)) {
        Write-Host "    ❌ $($nc.FirstName) $($nc.LastName)  [$($nc.Roles)]  — Missing: $($nc.MissingTraining)" -ForegroundColor Yellow
    }
}
Write-Host "  Warnings (no cred record): $warnCount" -ForegroundColor Yellow
if ($warningList.Count -gt 0) {
    foreach ($wv in ($warningList | Sort-Object LastName, FirstName)) {
        Write-Host "    ⚠️  $($wv.FirstName) $($wv.LastName)  [$($wv.Roles)]  — Issue: $($wv.Issue)" -ForegroundColor Yellow
    }
}
if ($errorCount -gt 0) { Write-Host "  Send errors     : $errorCount"           -ForegroundColor Red    }
Write-Host "  HTML previews   : $OutputFolder\"       -ForegroundColor Gray
Write-Host "  Compliant CSV   : $compliantCsvPath"    -ForegroundColor Green
Write-Host "  Non-Compliant CSV: $nonCompliantCsvPath" -ForegroundColor Yellow
if ($ExportReport) {
    Write-Host "  Full CSV report : $(Join-Path $OutputFolder 'ComplianceReport.csv')" -ForegroundColor Magenta
}

Write-Host ""
Write-Host "══════════════════════════════════════════════════" -ForegroundColor DarkCyan
Write-Host "  Run Complete"                                      -ForegroundColor White
Write-Host "══════════════════════════════════════════════════" -ForegroundColor DarkCyan
Write-Host "  Mode            : $modeLabel"
if ($SafetyOnly)         { Write-Host "  Safety Only     : Yes" -ForegroundColor Cyan }
if ($SendCompliantEmail) { Write-Host "  Compliant Emails: Yes" -ForegroundColor Cyan }
Write-Host "  Roles processed : $($selectedRoles -join ', ')"
Write-Host "  Emails sent     : $sentCount"           -ForegroundColor $(if ($sentCount -gt 0) {"Cyan"} else {"Gray"})
Write-Host "  Fully compliant : $($compliantList.Count)"      -ForegroundColor Green
Write-Host ""