# ==========================================
# Automatically Enable Litigation Hold for All Mailboxes
# ==========================================

$LogPath = "C:\Scripts\LitigationHold.log"
$Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

Add-Content $LogPath "`n==============================="
Add-Content $LogPath "Started: $Date"

Try {

    # Load Exchange Management Snap-in if not already loaded
    if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
    }

    # Get only mailboxes where Litigation Hold is disabled
    $Mailboxes = Get-Mailbox -ResultSize Unlimited |
                 Where-Object { -not $_.LitigationHoldEnabled }

    foreach ($Mailbox in $Mailboxes) {
        try {
            # Enable Litigation Hold for 1825 days (5 years)
            Set-Mailbox $Mailbox.Identity `
                -LitigationHoldEnabled $true `
                -LitigationHoldDuration 1825 `
                -LitigationHoldOwner "IT Department"

            Add-Content $LogPath "SUCCESS: Enabled for $($Mailbox.UserPrincipalName)"
        }
        catch {
            # Log mailbox-level errors
            Add-Content $LogPath "ERROR: $($Mailbox.UserPrincipalName) - $($_.Exception.Message)"
        }
    }

}
catch {
    # Log general script errors
    Add-Content $LogPath "ERROR (General): $($_.Exception.Message)"
}

$EndDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content $LogPath "Finished: $EndDate"
