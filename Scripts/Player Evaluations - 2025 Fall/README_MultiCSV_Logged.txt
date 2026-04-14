AYSO Player Evaluations Email Kit (with Logging & Archiving)
==============================================================

Included Files:
---------------
- Send-PlayerEvaluations-AllCSVs_Logged.ps1     → Batch runner with log/skip/archive support
- Send-PlayerEvaluations-SaveLocal.ps1          → Core script to send evaluations and save HTMLs
- Team_Roster_Report_12U_TEST.csv               → Example CSV #1 (Avalanche)
- Team_Roster_Report_12U_Bulldogs.csv           → Example CSV #2 (Bulldogs)
- Team_Roster_Report_12U_Tigers.csv             → Example CSV #3 (Tigers)
- README_MultiCSV_Logged.txt                    → This file

How to Use:
-----------
1. Open PowerShell.
2. Navigate to the unzipped folder.
3. Preview processing (no emails sent, logs actions):
   .\Send-PlayerEvaluations-AllCSVs_Logged.ps1 `
       -RawCSVPath "." `
       -GmailFrom "your@gmail.com" `
       -GmailPassword "your_app_password" `
       -DryRun

4. Send real emails but redirect all to yourself:
   .\Send-PlayerEvaluations-AllCSVs_Logged.ps1 `
       -RawCSVPath "." `
       -GmailFrom "your@gmail.com" `
       -GmailPassword "your_app_password" `
       -RedirectTo "your@gmail.com"

Features:
---------
✔️ Skips previously processed CSVs using 'processed_files.log'
✔️ Archives finished files to 'archived/' folder
✔️ Saves a copy of each email as HTML in 'EvaluationEmailOutput/'

Requirements:
-------------
- PowerShell 5+ or PowerShell Core
- Gmail account with App Password (if 2FA is on)
