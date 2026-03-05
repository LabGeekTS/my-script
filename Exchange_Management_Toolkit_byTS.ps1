<#
.SYNOPSIS
    Exchange Server Administrative Toolkit
.DESCRIPTION
    A centralized, guide-based script for routine Exchange management.
    Includes Tracking, Distribution Groups, and Phishing Cleanup.
#>

function Initialize-Environment {
    if (!(Test-Path "C:\Temp")) {
        New-Item -Path "C:\Temp" -ItemType Directory | Out-Null
        Write-Host "[INIT] Created C:\Temp for exports." -ForegroundColor Gray
    }
}

function Show-Header {
    Clear-Host
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host "         EXCHANGE SERVER ADMINISTRATIVE TOOLKIT           " -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan
}

Initialize-Environment

# We add a Label (:MainMenu) to the loop so we can break out of it properly
:MainMenu while($true) {
    Show-Header
    Write-Host "1. Message Tracking (Search by Sender/Recipient)" -ForegroundColor White
    Write-Host "2. Export Distribution Group Members to CSV" -ForegroundColor White
    Write-Host "3. Bulk ADD Members to Distribution Group (CSV)" -ForegroundColor Green
    Write-Host "4. Bulk REMOVE Members from Distribution Group (CSV)" -ForegroundColor Red
    Write-Host "5. Search and Delete Mail by SUBJECT (Phishing Cleanup)" -ForegroundColor Magenta
    Write-Host "6. Search and Delete Mail by SENDER (Phishing Cleanup)" -ForegroundColor Magenta
    Write-Host "Q. Quit / Exit" -ForegroundColor Red
    Write-Host "----------------------------------------------------------"
    
    $choice = Read-Host "Please select an option (1-6 or Q)"

    switch($choice) {
        1 {
            Write-Host "`n[ GUIDE: MESSAGE TRACKING ]" -ForegroundColor Yellow
            $sender = Read-Host "Enter Sender Address"
            $recipient = Read-Host "Enter Recipient Address"
            $exportPath = "C:\Temp\TrackingLog_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
            Get-MessageTrackingLog -Sender $sender -Recipient $recipient -ResultSize Unlimited | 
            Select-Object Timestamp, EventId, Source, Sender, @{Name='Recipients';Expression={$_.Recipients -join ";"}}, MessageSubject | 
            Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
            Write-Host "Results saved to: $exportPath" -ForegroundColor Green
            Pause
        }

        2 {
            Write-Host "`n[ GUIDE: GROUP MEMBER EXPORT ]" -ForegroundColor Yellow
            $groupName = Read-Host "Enter Distribution Group Identity"
            $exportPath = "C:\Temp\Members_$($groupName -replace '[@.]', '_').csv"
            Get-DistributionGroupMember -Identity $groupName -ResultSize Unlimited | 
            Select-Object PrimarySmtpAddress | 
            Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
            Write-Host "Exported to: $exportPath" -ForegroundColor Green
            Pause
        }

        3 {
            Write-Host "`n[ GUIDE: BULK ADD MEMBERS ]" -ForegroundColor Green
            $groupName = Read-Host "Enter Target Group"
            $csvPath = Read-Host "Enter CSV Path (header: 'Email')"
            if (Test-Path $csvPath) {
                Import-Csv $csvPath | ForEach-Object { 
                    Add-DistributionGroupMember -Identity $groupName -Member $_.Email -ErrorAction SilentlyContinue 
                    Write-Host "Processing: $($_.Email)" -ForegroundColor Gray
                }
                Write-Host "Bulk Add Completed." -ForegroundColor Green
            } else { Write-Host "File not found!" -ForegroundColor Red }
            Pause
        }

        4 {
            Write-Host "`n[ GUIDE: BULK REMOVE MEMBERS ]" -ForegroundColor Red
            $groupName = Read-Host "Enter Target Group"
            $csvPath = Read-Host "Enter CSV Path (header: 'Email')"
            if (Test-Path $csvPath) {
                Import-Csv $csvPath | ForEach-Object { 
                    Remove-DistributionGroupMember -Identity $groupName -Member $_.Email -Confirm:$false -ErrorAction SilentlyContinue 
                    Write-Host "Processing: $($_.Email)" -ForegroundColor Gray
                }
                Write-Host "Bulk Remove Completed." -ForegroundColor Green
            } else { Write-Host "File not found!" -ForegroundColor Red }
            Pause
        }

        5 {
            Write-Host "`n[ GUIDE: DELETE BY SUBJECT ]" -ForegroundColor Magenta
            $csvPath = Read-Host "Enter Path to CSV list"
            $subjectInput = Read-Host "Enter Subject to search"
            $subjectQuery = "Subject:""$subjectInput"""
            $adminMailbox = Read-Host "Enter Admin Mailbox for logs"
            $logFolder = "PhishingSubjectLogs"
            $exportPath = "C:\Temp\Deletion_Subject_Report_$(Get-Date -Format 'ddMMyy_HHmm').csv"

            if (!(Test-Path $csvPath)) { Write-Host "CSV not found!" -ForegroundColor Red; Pause; continue }

            $list = Import-Csv $csvPath
            $foundList = @()
            $reportData = @()

            Write-Host "Step 1: Scanning mailboxes by SUBJECT..." -ForegroundColor Cyan
            foreach ($row in $list) {
                $email = $row.Email
                try {
                    $check = Search-Mailbox -Identity $email -SearchQuery $subjectQuery -LogOnly -LogLevel Basic -TargetMailbox $adminMailbox -TargetFolder $logFolder -ErrorAction Stop
                    if ($check.ResultItemsCount -gt 0) {
                        Write-Host "MATCH FOUND: $email ($($check.ResultItemsCount) items)" -ForegroundColor Yellow
                        $foundList += [PSCustomObject]@{ Mailbox = $email; ItemsFound = $check.ResultItemsCount }
                    }
                } catch { Write-Host "ERROR scanning $email" -ForegroundColor Red }
            }

            if ($foundList.Count -gt 0) {
                Write-Host "`nFound $($foundList.Count) mailboxes. Type 'OK' to proceed with deletion." -ForegroundColor White
                $confirm = Read-Host "Confirmation"
                if ($confirm -eq "OK" -or $confirm -eq "ok") {
                    foreach ($item in $foundList) {
                        $user = $item.Mailbox
                        try {
                            Search-Mailbox -Identity $user -SearchQuery $subjectQuery -DeleteContent -Force -ErrorAction Stop | Out-Null
                            Write-Host "$user : Deleted." -ForegroundColor Gray
                            $reportData += [PSCustomObject]@{ Mailbox = $user; ItemsDeleted = $item.ItemsFound; Status = "Deleted"; DateExecuted = Get-Date -Format "yyyy-MM-dd HH:mm:ss" }
                        } catch {
                            $reportData += [PSCustomObject]@{ Mailbox = $user; ItemsDeleted = 0; Status = "Failed"; DateExecuted = Get-Date -Format "yyyy-MM-dd HH:mm:ss" }
                        }
                    }
                    $reportData | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
                    Write-Host "SUCCESS! Report saved to: $exportPath" -ForegroundColor Green
                } else { Write-Host "Aborted." -ForegroundColor DarkYellow }
            } else { Write-Host "No matches found." -ForegroundColor White }
            Pause
        }

        6 {
            Write-Host "`n[ GUIDE: DELETE BY SENDER ]" -ForegroundColor Magenta
            $csvPath = Read-Host "Enter Path to CSV list"
            $senderAddr = Read-Host "Enter Sender Address to delete"
            $senderQuery = "From:""$senderAddr"""
            $adminMailbox = Read-Host "Enter Admin Mailbox for logs"
            $logFolder = "PhishingSenderLogs"
            $exportPath = "C:\Temp\Deletion_Sender_Report_$(Get-Date -Format 'ddMMyy_HHmm').csv"

            if (!(Test-Path $csvPath)) { Write-Host "CSV not found!" -ForegroundColor Red; Pause; continue }

            $list = Import-Csv $csvPath
            $foundList = @()
            $reportData = @()

            Write-Host "Step 1: Scanning mailboxes by SENDER..." -ForegroundColor Cyan
            foreach ($row in $list) {
                $email = $row.Email
                try {
                    $check = Search-Mailbox -Identity $email -SearchQuery $senderQuery -LogOnly -LogLevel Basic -TargetMailbox $adminMailbox -TargetFolder $logFolder -ErrorAction Stop
                    if ($check.ResultItemsCount -gt 0) {
                        Write-Host "MATCH FOUND: $email ($($check.ResultItemsCount) items)" -ForegroundColor Yellow
                        $foundList += [PSCustomObject]@{ Mailbox = $email; ItemsFound = $check.ResultItemsCount }
                    }
                } catch { Write-Host "ERROR scanning $email" -ForegroundColor Red }
            }

            if ($foundList.Count -gt 0) {
                Write-Host "`nFound $($foundList.Count) mailboxes. Type 'OK' to proceed with deletion." -ForegroundColor White
                $confirm = Read-Host "Confirmation"
                if ($confirm -eq "OK" -or $confirm -eq "ok") {
                    foreach ($item in $foundList) {
                        $user = $item.Mailbox
                        try {
                            Search-Mailbox -Identity $user -SearchQuery $senderQuery -DeleteContent -Force -ErrorAction Stop | Out-Null
                            Write-Host "$user : Deleted." -ForegroundColor Gray
                            $reportData += [PSCustomObject]@{ Mailbox = $user; ItemsDeleted = $item.ItemsFound; Status = "Deleted"; DateExecuted = Get-Date -Format "yyyy-MM-dd HH:mm:ss" }
                        } catch {
                            $reportData += [PSCustomObject]@{ Mailbox = $user; ItemsDeleted = 0; Status = "Failed"; DateExecuted = Get-Date -Format "yyyy-MM-dd HH:mm:ss" }
                        }
                    }
                    $reportData | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
                    Write-Host "SUCCESS! Report saved to: $exportPath" -ForegroundColor Green
                } else { Write-Host "Aborted." -ForegroundColor DarkYellow }
            } else { Write-Host "No matches found." -ForegroundColor White }
            Pause
        }

        "q" { 
            Write-Host "`nExiting Administrative Toolkit. Goodbye!" -ForegroundColor Cyan
            # This now breaks the loop labeled MainMenu
            break MainMenu 
        }

        default {
            Write-Host "Invalid selection. Please choose 1-6 or Q." -ForegroundColor Red
            Pause
        }
    }
}