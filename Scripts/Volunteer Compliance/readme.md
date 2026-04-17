# Outputs html and compliance report
.\Send-TrainingReminders.ps1 -DryRun -ExportReport -Roles "HC,AC,YR,Ref,TM,FS"  

# Sends Emails To asyo138.ca@gmail.com for testing emails
.\Send-TrainingReminders.ps1 -ExportReport -Roles "HC,AC,YR,Ref,TM,FS" -TestMode

# Actually sends emails with compliance report and html and cc'd
.\Send-TrainingReminders.ps1 -ExportReport -Roles "HC,AC,YR,Ref,TM,FS" -CC "ayso138.ca@gmail.com"

# Failed Emails
.\Send-TrainingReminders.ps1 -ExportReport -Roles "HC,AC,YR,Ref,TM,FS" -CC "ayso138.ca@gmail.com" -OnlyList "FailedEmails.txt" -DelayMs 2000

# Next Season
# First run — sends everyone, logs who was sent
.\Send-TrainingReminders.ps1 -Roles All -ExportReport -CC "ayso138.ca@gmail.com" -SafetOnly

# If it breaks mid-run, re-run with -SkipAlreadySent to pick up where it left off
.\Send-TrainingReminders.ps1 -Roles All -ExportReport -CC "ayso138.ca@gmail.com" -SafetyOnly -SkipAlreadySent