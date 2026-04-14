<#
.SYNOPSIS
📇 Creates or updates iCloud contacts for AYSO coaches from a team roster CSV.

.DESCRIPTION
- Adds Head Coach and Assistant Coach contacts to iCloud Contacts on macOS.
- Adds contacts to a dynamic group: "AYSO - YEAR SEASON - Coaches - DIVISION"
- Appends team info to the Notes field (deduplicated).
- Updates existing contacts if first+last name match.
- Creates new contacts if no match is found.
- Prevents duplicate contacts by checking before adding.
- Activates the contact in Contacts.app after update or creation.
- Supports dry run mode to preview without making changes.
- Logs created/updated contacts with timestamp to a single rolling log file.
#>

param (
    [string]$RawCSVPath = "Team_Roster_Report_6U_Coed.csv",
    [string]$Year = "2025",
    [string]$Season = "F",
    [bool]$DryRun = $true
)

# Use one rolling log file
$logFile = "Contacts_Processed.log"
if (-not (Test-Path $logFile)) {
    "Timestamp,Name,Email,Status,Group,Note" | Out-File -FilePath $logFile
}

# Extract division from filename (everything after Team_Roster_Report_ until .csv)
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($RawCSVPath)
if ($baseName -like "Team_Roster_Report_*") {
    $division = $baseName.Substring("Team_Roster_Report_".Length)
}
else {
    $division = $baseName
}

# Replace underscores with spaces for readability
$division = $division -replace "_", " "

$GroupName = "AYSO - $Year $Season - Coaches - $division"

Write-Host "📄 Processing: $RawCSVPath"

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

    # Safely extract team name (avoid array issues)
    $teamLine = ($section | Where-Object { $_ -like "Team Name:*" } | Select-Object -First 1)
    if ($null -ne $teamLine) {
        $teamName = $teamLine -replace "Team Name: ", ""
        $teamName = $teamName -replace ',', ''
        $teamName = $teamName.Trim()
    }
    else {
        $teamName = ""
    }

    $coachStart = ($section | Select-String "Team Personnel").LineNumber
    if ($null -eq $coachStart) {
        if ([string]::IsNullOrWhiteSpace($teamName) -eq $false) {
            Write-Warning "⚠️ No 'Team Personnel' section found for team: $teamName — skipping"
        }
        continue
    }

    $coachLines = $section[($coachStart + 1)..($section.Count - 1)]
    $coaches = @()

    foreach ($line in $coachLines) {
        $cols = $line -split ","
        if ($cols.Length -lt 5) { continue }

        $parsed = [PSCustomObject]@{
            Role      = $cols[1].Trim()
            FirstName = $cols[2].Trim()
            LastName  = $cols[3].Trim()
            Email     = $cols[4].Trim()
            Phone     = if ($cols.Length -gt 5) { $cols[5].Trim() } else { "" }
        }

        if (($parsed.Role -eq "Head Coach" -or $parsed.Role -eq "Assistant Coach") -and $parsed.Email -match "@") {
            $coaches += $parsed
        }
    }

    if ($coaches.Count -eq 0) {
        Write-Warning "⚠️ No coaches for team: $teamName — skipping"
        continue
    }

    foreach ($coach in $coaches) {
        $FirstName = $coach.FirstName
        $CoachRoleShort = $coach.Role -replace 'Head Coach', 'HC' -replace 'Assistant Coach', 'AC'
        $LastNameFormatted = "($teamName $CoachRoleShort) $($coach.LastName)"
        $Email = $coach.Email
        $Phone = $coach.Phone
        $Notes = "AYSO - $Year $Season – $teamName – $($coach.Role)"

        $appleScript = @"
tell application "Contacts"
	if not (exists group "$GroupName") then
		make new group with properties {name:"$GroupName"}
	end if

	set foundContact to missing value

	-- Lookup requires first and last name only
	try
		set foundContact to first person whose (first name is "$FirstName" and last name is "$LastNameFormatted")
	end try

	if foundContact is not missing value then
		-- Update existing
		set alreadyHasNote to false
		if note of foundContact contains "$Notes" then
			set alreadyHasNote to true
		end if

		set first name of foundContact to "$FirstName"
		set last name of foundContact to "$LastNameFormatted"

		if not alreadyHasNote then
			if note of foundContact is not "" then
				set note of foundContact to (note of foundContact & return & "$Notes")
			else
				set note of foundContact to "$Notes"
			end if
		end if

		add foundContact to group "$GroupName"
	else
		-- Create new contact
		set myCard to make new person with properties {first name:"$FirstName", last name:"$LastNameFormatted", note:"$Notes"}
		tell myCard
			if "$Email" is not "" then
				make new email at end of emails with properties {label:"home", value:"$Email"}
			end if
			if "$Phone" is not "" then
				make new phone at end of phones with properties {label:"mobile", value:"$Phone"}
			end if
		end tell
		add myCard to group "$GroupName"
		set foundContact to myCard
		save
	end if
end tell

tell application "Contacts"
	if foundContact is not missing value then
		activate
		set selection to foundContact
	end if
end tell
"@

        if (-not $DryRun) {
            osascript -e $appleScript
            $status = if ($?) { "Created/Updated" } else { "Error" }
            $timestampNow = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            "$timestampNow,$FirstName $LastNameFormatted,$Email,$status,$GroupName,""$Notes""" | Out-File -FilePath $logFile -Append
        }
        else {
            Write-Host "🧪 Dry run — would create/update: $FirstName $LastNameFormatted <$Email> ($Notes)"
        }
    }
}

if (-not $DryRun) {
    Write-Host "`n📄 Log updated: $logFile"
}