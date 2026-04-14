<#
.SYNOPSIS
📋 Sends player evaluation emails to AYSO 12U coaches based on a raw Team Roster CSV.

.DESCRIPTION
🔁 Parses a multi-team CSV export, extracts players and coaches,
filters Head/Assistant Coaches, and sends a styled HTML email with player info.

.PARAMETER RawCSVPath
Path to the full team roster CSV file.

.PARAMETER GmailFrom
Gmail address to send from.

.PARAMETER GmailPassword
Gmail app password.

.PARAMETER DryRun
If specified, does not send emails—just saves HTML previews and prints summary.

.PARAMETER RedirectTo
If specified, sends all emails to this address instead of the real coaches (for testing).

.EXAMPLE
.\Send-PlayerEvaluations-SaveLocal.ps1 -RawCSVPath "Team_Roster_Report_12U.csv" -GmailFrom "your@gmail.com" -GmailPassword "app-password" -DryRun
#>

param (
    [string]$RawCSVPath = "Team_Roster_Report_TEST.csv",
    [string]$GmailFrom = "ayso138.ca@gmail.com",
    [string]$GmailPassword = "rovn yjfh ansj popl",
    [switch]$DryRun = $true,
    [string]$RedirectTo = "john@rennemeyer.com"
)

# Create output folder
$outputFolder = "EvaluationEmailOutput"
if (!(Test-Path -Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

# Read the raw CSV content
$lines = Get-Content $RawCSVPath
$teamSections = @()
$current = @()

foreach ($line in $lines) {
    if ($line -like "Team Name:*") {
        if ($current.Count -gt 0) { $teamSections += , @($current); $current = @() }
    }
    $current += $line
}
if ($current.Count -gt 0) { $teamSections += , @($current) }

foreach ($section in $teamSections) {
    $teamName = ($section | Where-Object { $_ -like "Team Name:*" }) -replace "Team Name: ", ""
    $teamName = $teamName -replace ',', ''
    $safeTeamName = ($teamName -replace '[^a-zA-Z0-9]', '_')

    $playerStart = ($section | Select-String "Team Players").LineNumber + 1
    $coachStart = ($section | Select-String "Team Personnel").LineNumber
    $playerLines = $section[$playerStart..($coachStart - 1)] | Where-Object { $_ -match "^\d+,.+,.+" }

    $players = @()
    foreach ($line in $playerLines) {
        $cols = $line -split ","
        if ($cols.Length -ge 3) {
            $players += [PSCustomObject]@{
                'First Name' = $cols[1].Trim()
                'Last Name'  = $cols[2].Trim()
            }
        }
    }

    $coachLines = $section[($coachStart + 1)..($section.Count - 1)]
    $headCoach = $null
    $assistantCoach = $null

    foreach ($line in $coachLines) {
        $cols = $line -split ","
        if ($cols.Length -lt 5) { continue }
        if ($cols[1] -eq "Head Coach" -and $cols[4] -match "@") {
            $headCoach = [PSCustomObject]@{
                Role = $cols[1]; FirstName = $cols[2]; LastName = $cols[3]; Email = $cols[4]
            }
        }
        elseif ($cols[1] -eq "Assistant Coach" -and $cols[4] -match "@") {
            $assistantCoach = [PSCustomObject]@{
                Role = $cols[1]; FirstName = $cols[2]; LastName = $cols[3]; Email = $cols[4]
            }
        }
    }

    if (-not $headCoach) {
        if ($assistantCoach) {
            $headCoach = $assistantCoach
            $assistantCoach = $null    
        }
        else {
            Write-Warning "⚠️ No head coach or assistant coach for team: $teamName — skipping email"
            continue
        }
    }

    # Build player rows
    $playerRows = ""
    foreach ($p in $players) {
        $playerRows += "<tr><td>$($p.'First Name')</td><td>$($p.'Last Name')</td><td></td></tr>`n"
    }

    # Build styled HTML
    $htmlBody = @"
<html>
<head>
  <style>
    body {
      font-family: Segoe UI, Arial, sans-serif;
      font-size: 14px;
      color: #333;
      line-height: 1.6;
    }
    table {
      border-collapse: collapse;
      width: 100%;
      margin-top: 10px;
    }
    th, td {
      border: 1px solid #ccc;
      padding: 6px 10px;
      text-align: left;
    }
    th {
      background-color: #f4f4f4;
    }
    .rating-guide {
      margin-top: 10px;
    }
  </style>
</head>
<body>
  <p>Hi Coach $($headCoach.LastName),</p>

  <p>Thank you for your help coaching this fall! We're gathering evaluations for the players that were on your team, <strong>$teamName</strong>, to help with balanced team formation this spring.</p>

  <p>Please review the roster below and reply to this email by November 1st, 2025, with a rating for each player in the "Evaluation (1-5)" column based on how they rank against all players in your division (not just their team). Feel free to include any notes you may have on the player. Also, please mark if a player is a top scorer on the team with "TS".</p>
  <p><strong>How to rate players:</strong></p>

  <div class="rating-guide">
    <p><strong>5</strong> - High impact player; can carry a team; excellent individual and team skills; leader; can change the way the game goes if they are absent.</p>
    <p><strong>4</strong> - Strong player; good individual and team skills; excels at one or more positions (e.g., goalkeeper).</p>
    <p><strong>3</strong> - Average player; has basic skills and game understanding; generally neutral impact on play.</p>
    <p><strong>2</strong> - Below average player; lacks some basic skills or field awareness.</p>
    <p><strong>1</strong> - New or disruptive player; little experience or understanding of the game.</p>
    <p>You may use <strong>0.5</strong> increments if needed for better differentiation (e.g., 3.5).</p>
  </div>

  <table>
    <thead>
      <tr>
        <th>First Name</th>
        <th>Last Name</th>
        <th>Evaluation (1-5), Top Scorer, Notes</th>        
      </tr>
    </thead>
    <tbody>
      $playerRows
    </tbody>
  </table>

  <p>If you have any questions or feedback, feel free to reply. We really appreciate your support!</p>

  <p><strong>John Rennemeyer</strong><br>
     Coach Administrator<br>
     <a href="https://www.ayso138.org/">AYSO Region 138</a></p>
</body>
</html>
"@

    # Save to local file
    $htmlPath = Join-Path $outputFolder "$safeTeamName.html"
    $htmlBody | Out-File -FilePath $htmlPath -Encoding utf8

    # Determine recipients
    $to = if ($RedirectTo) { $RedirectTo } else { $headCoach.Email }
    $cc = if ($RedirectTo) { @() } elseif ($assistantCoach) { $assistantCoach.Email } else { @() }

    if ($DryRun) {
        Write-Host "`n📝 DRY RUN: Would send email for $teamName"
        Write-Host "To: $to"
        if ($cc) { Write-Host "CC: $cc" }
        Write-Host "Subject: 12U Player Evaluations - $teamName"
        Write-Host "Email saved to: $htmlPath"
    }
    else {
        # Send email (disabled by default; uncomment to enable)
        $params = @{
            From       = "AYSO 138 Coach Admin <$GmailFrom>"
            To         = $to
            Subject    = "$teamName Player Evaluations"
            Body       = $htmlBody
            BodyAsHtml = $true
            SmtpServer = "smtp.gmail.com"
            Port       = 587
            UseSsl     = $true
            Credential = New-Object System.Management.Automation.PSCredential($GmailFrom, (ConvertTo-SecureString $GmailPassword -AsPlainText -Force))
        }

        if ($cc -and $cc.Count -gt 0) {
            $params.Cc = $cc
        }

        Send-MKMailMessage @params
    
        Write-Host "📧 Sent to $to (cc: $cc) and saved to $htmlPath"
    }

}
