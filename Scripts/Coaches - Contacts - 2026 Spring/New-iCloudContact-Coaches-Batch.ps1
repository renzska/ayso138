<#
.SYNOPSIS
📦 Batch import AYSO coach contacts to iCloud from one or more CSV files using AppleScript.

.DESCRIPTION
- Skips previously processed CSVs using a log file (`processed_files.log`)
- Calls `New-iCloudContact-Coaches-Single.ps1` per CSV
- Dynamically determines Year, Season, and Division from context
- DryRun mode enabled by default unless overridden
- Designed to process a single file or an entire folder

.PARAMETER RawCSVPath
A path to either a single CSV file or a folder containing multiple team roster CSV files.

.PARAMETER Year
The season year for the group name and contact notes (e.g., 2025)

.PARAMETER Season
The season name (e.g., Spring, Fall) for the group and notes

.PARAMETER DryRun
If true, does not create/update contacts, but shows what would be done

.EXAMPLE
# Process one file for Spring (S)
.\New-iCloudContact-Coaches-Batch.ps1 -RawCSVPath "Team_Roster_Report_6U.csv" -Year 2026 -Season "S"

.EXAMPLE
# Process all .csv files in a folder (dry run by default)
.\New-iCloudContact-Coaches-Batch.ps1 -RawCSVPath "." -Year 2026 -Season "S"

.EXAMPLE
# Process all with real changes
.\New-iCloudContact-Coaches-Batch.ps1 -RawCSVPath "." -Year 2026 -Season "S" -DryRun:$false
#>

param (
    [string]$RawCSVPath = ".",
    [string]$Year = "2026",
    [string]$Season = "S",
    [bool]$DryRun = $true
)

$logFile = "processed_files.log"
if (!(Test-Path $logFile)) {
    New-Item -ItemType File -Path $logFile | Out-Null
}

function Has-BeenProcessed {
    param ([string]$fileName)
    return (Select-String -Path $logFile -Pattern ([regex]::Escape($fileName)) -Quiet)
}

function Mark-Processed {
    param ([string]$fileName)
    Add-Content -Path $logFile -Value $fileName
}

function Process-CsvFile {
    param ([System.IO.FileInfo]$file)

    if (Has-BeenProcessed $file.Name) {
        Write-Host "⏭️ Skipping (already processed): $($file.Name)" -ForegroundColor Yellow
        return
    }

    Write-Host "`n📄 Processing: $($file.Name)" -ForegroundColor Cyan

    & ".\New-iCloudContact-Coaches-Single.ps1" `
        -RawCSVPath $file.FullName `
        -Year $Year `
        -Season $Season `
        -DryRun:$DryRun

    if (-not $DryRun) {
        Mark-Processed $file.Name
    }
}

if (Test-Path $RawCSVPath -PathType Leaf) {
    # Single CSV file
    $fileInfo = Get-Item $RawCSVPath
    Process-CsvFile -file $fileInfo
}
elseif (Test-Path $RawCSVPath -PathType Container) {
    # Folder of .csv files
    $csvFiles = Get-ChildItem -Path $RawCSVPath -Filter *.csv
    foreach ($file in $csvFiles) {
        Process-CsvFile -file $file
    }
}
else {
    Write-Error "❌ Provided path '$RawCSVPath' is not a valid file or folder."
}