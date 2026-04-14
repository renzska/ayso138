<#
.SYNOPSIS
📋 Processes one or more AYSO player roster CSVs and sends evaluation emails to coaches.

.DESCRIPTION
- Skips previously processed CSVs using a log.
- Archives completed CSVs to a subfolder.

.EXAMPLE
# Process one file
.\Send-PlayerEvaluations-AllCSVs_Logged.ps1 -RawCSVPath "Team_Roster.csv" -GmailFrom "me@gmail.com" -GmailPassword "app-pass" -DryRun

.EXAMPLE
# Process all CSVs in a folder
.\Send-PlayerEvaluations-AllCSVs_Logged.ps1 -RawCSVPath "." -GmailFrom "me@gmail.com" -GmailPassword "app-pass" -RedirectTo "me@gmail.com"
#>

param (
    [string]$RawCSVPath = ".",
    [string]$GmailFrom = "ayso138.ca@gmail.com",
    [string]$GmailPassword = "rovn yjfh ansj popl",
    [switch]$DryRun = $false,
    [string]$RedirectTo = ""
)

$logFile = "processed_files.log"

if (!(Test-Path $logFile)) {
    New-Item $logFile -ItemType File | Out-Null
}

function Has-BeenProcessed {
    param ($fileName)
    # FIX: ensure -Pattern is explicitly set and value is evaluated, avoiding positional binding
    return (Select-String -Path $logFile -Pattern ([regex]::Escape($fileName)) -Quiet)
}

function Mark-Processed {
    param ($fileName)
    Add-Content -Path $logFile -Value $fileName
}

function Process-CsvFile {
    param ([System.IO.FileInfo]$file)

    if (Has-BeenProcessed $file.Name) {
        Write-Host "⏭️ Skipping (already processed): $($file.Name)" -ForegroundColor Yellow
        return
    }

    Write-Host "`n📄 Processing: $($file.Name)" -ForegroundColor Cyan

    & ".\Send-PlayerEvaluations-SaveLocal.ps1" `
        -RawCSVPath $file.FullName `
        -GmailFrom $GmailFrom `
        -GmailPassword $GmailPassword `
        -RedirectTo $RedirectTo `
        -DryRun:($DryRun.IsPresent)

    Mark-Processed $file.Name
}

if (Test-Path $RawCSVPath -PathType Leaf) {
    # Single file
    $fileInfo = Get-Item $RawCSVPath
    Process-CsvFile -file $fileInfo
}
elseif (Test-Path $RawCSVPath -PathType Container) {
    # Folder: Process all .csv files
    $csvFiles = Get-ChildItem -Path $RawCSVPath -Filter *.csv
    foreach ($file in $csvFiles) {
        Process-CsvFile -file $file
    }
}
else {
    Write-Error "❌ Provided path '$RawCSVPath' is not a valid file or folder."
}
