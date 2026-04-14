<#
.SYNOPSIS
📇 Creates or updates iCloud contacts for AYSO parents from a team roster CSV.

.DESCRIPTION
- Adds parent contacts from the "Team Players" section.
- Adds each contact to both:
  • AYSO - 2025 F - Parents - [Division]
  • AYSO - 2025 SEASON - All Parents
- Formats name and notes with child and team info.
- Prevents duplicates. Supports dry-run mode.
#>

param (
    [string]$RawCSVPath = "Team_Roster_Report_6U.csv",
    [string]$Year = "2025",
    [string]$Season = "F",
    [bool]$DryRun = $true
)

$logFile = "Contacts_Processed.log"
if (-not (Test-Path $logFile)) {
    "Timestamp,Name,Email,Status,Group,Note" | Out-File -FilePath $logFile
}

# Extract division from filename
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($RawCSVPath)
$division = if ($baseName -like "Team_Roster_Report_*") {
    $baseName.Substring("Team_Roster_Report_".Length) -replace "_", " "
} else {
    $baseName -replace "_", " "
}

$DivisionGroup = "AYSO - $Year $Season - Parents - $division"
$MasterGroup   = "AYSO - $Year $Season - Parents All"

Write-Host "📄 Processing: $RawCSVPath"

# Read raw content
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

    $teamLine = ($section | Where-Object { $_ -like "Team Name:*" } | Select-Object -First 1)
    $teamName = if ($null -ne $teamLine) {
        ($teamLine -replace "Team Name: ", "").Trim().Replace(",", "")
    } else { "" }

    $playerStart = ($section | Select-String "Team Players").LineNumber
    $coachStart = ($section | Select-String "Team Personnel").LineNumber

    if ($null -eq $playerStart -or $null -eq $coachStart) {
        Write-Warning "⚠️ Skipping team: $teamName — missing player/personnel section"
        continue
    }

    $playerLines = $section[($playerStart + 1)..($coachStart - 1)]
    foreach ($line in $playerLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $cols = $line -split ","
        if ($cols.Length -lt 7) { continue }

        $childFirst = $cols[1].Trim()
        $childLast  = $cols[2].Trim()
        $parentFirst = $cols[3].Trim()
        $parentLast  = $cols[4].Trim()
        $email = $cols[5].Trim()
        $phone = $cols[6].Trim()

        if ($email -notmatch "@") { continue }

        $lastNameFormatted = if ($childLast -ne $parentLast) {
            "($childFirst $childLast - $teamName) $parentLast"
        } else {
            "($childFirst - $teamName) $parentLast"
        }

        $note = "$childFirst $childLast – $Year $Season – $teamName"

        $appleScript = @"
tell application "Contacts"
	-- Ensure both groups exist
	if not (exists group "$DivisionGroup") then
		make new group with properties {name:"$DivisionGroup"}
	end if
	if not (exists group "$MasterGroup") then
		make new group with properties {name:"$MasterGroup"}
	end if

	set foundContact to missing value
	try
		set foundContact to first person whose (first name is "$parentFirst" and last name is "$lastNameFormatted")
	end try

	if foundContact is not missing value then
		set alreadyHasNote to false
		if note of foundContact contains "$note" then set alreadyHasNote to true

		set first name of foundContact to "$parentFirst"
		set last name of foundContact to "$lastNameFormatted"

		if not alreadyHasNote then
			if note of foundContact is not "" then
				set note of foundContact to (note of foundContact & return & "$note")
			else
				set note of foundContact to "$note"
			end if
		end if

		add foundContact to group "$DivisionGroup"
		add foundContact to group "$MasterGroup"
	else
		set myCard to make new person with properties {first name:"$parentFirst", last name:"$lastNameFormatted", note:"$note"}
		tell myCard
			if "$email" is not "" then
				make new email at end of emails with properties {label:"home", value:"$email"}
			end if
			if "$phone" is not "" then
				make new phone at end of phones with properties {label:"mobile", value:"$phone"}
			end if
		end tell
		add myCard to group "$DivisionGroup"
		add myCard to group "$MasterGroup"
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
            "$timestampNow,$parentFirst $lastNameFormatted,$email,$status,$DivisionGroup + $MasterGroup,""$note""" | Out-File -FilePath $logFile -Append
            Write-Host "✅ Added: $parentFirst $lastNameFormatted <$email> ($note)"
        }
        else {
            Write-Host "🧪 Dry run — would create/update: $parentFirst $lastNameFormatted <$email> ($note)"
        }
    }
}

if (-not $DryRun) {
    Write-Host "`n📄 Log updated: $logFile"
}