<#
.SYNOPSIS
📦 Batch import AYSO parent contacts to iCloud from one or more CSV files using AppleScript.

.DESCRIPTION
- Calls `New-iCloudContact-Parents-Single.ps1` per CSV file.
- Each parent contact is added to:
  • Division group: "AYSO - YEAR SEASON - Parents - [Division]"
  • Master group:   "AYSO - YEAR SEASON - All Parents"
- Skips previously processed CSVs using a log file (`processed_files.log`)
- Dynamically determines Year, Season, and Division from file name
- DryRun mode enabled by default unless overridden
- Designed to process a single file or an entire folder

.PARAMETER RawCSVPath
A path to either a single CSV file or a folder containing multiple team roster CSV files.

.PARAMETER Year
The season year for the group name and contact notes (e.g., 2025)

.PARAMETER Season
The season name/letter (e.g., F for Fall, S for Spring) for the group and notes

.PARAMETER DryRun
If true, does not create/update contacts, but shows what would be done

.EXAMPLE
# Process one file for Spring (S)
.\New-iCloudContact-Parents-Batch.ps1 -RawCSVPath "Team_Roster_Report_6U.csv" -Year 2026 -Season "S"

.EXAMPLE
# Process all .csv files in a folder (dry run by default)
.\New-iCloudContact-Parents-Batch.ps1 -RawCSVPath "." -Year 2026 -Season "S"

.EXAMPLE
# Process all with real changes
.\New-iCloudContact-Parents-Batch.ps1 -RawCSVPath "." -Year 2026 -Season "S" -DryRun:$false
#>

param (
    [string]$RawCSVPath = ".",
    [string]$Year = "2026",
    [string]$Season = "S",
    [bool]$DryRun = $true
)

$logFile = "processed_files.log"
if (-not (Test-Path $logFile)) {
    "" | Out-File -FilePath $logFile
}
$alreadyProcessed = Get-Content $logFile

# Get file list
if (Test-Path $RawCSVPath -PathType Container) {
    $csvFiles = Get-ChildItem -Path $RawCSVPath -Recurse -Include *.csv
}
else {
    $csvFiles = @((Get-Item -Path $RawCSVPath))
}

foreach ($file in $csvFiles) {
    $fullPath = $file.FullName
    $filename = $file.Name

    if ($alreadyProcessed -contains $filename) {
        Write-Host "⏩ Skipping already processed file: $filename"
        continue
    }

    Write-Host "📋 Processing: $filename"
    & pwsh ./New-iCloudContact-Parents-Single.ps1 -RawCSVPath $fullPath -Year $Year -Season $Season -DryRun:$DryRun

    if (-not $DryRun) {
        $filename | Out-File -FilePath $logFile -Append
    }
}