<# PARAMS
.SYNOPSIS
    Collects information on Shared and Room Mailboxes in a Microsoft 365 tenant filtered by a specified domain.
.DESCRIPTION
    This script connects to Exchange Online, retrieves all Shared and Room Mailboxes, filters them based on the provided domain, and collects details such as Display Name, Primary Email, Mailbox Type, Mailbox Size in MB, Aliases, and LegacyExchangeDN. The collected data is then exported to a CSV file for further analysis.
.PARAMETER AdminUPN
    The User Principal Name (UPN) of the admin account used to connect to Exchange Online.
.PARAMETER DomainFilter
    The domain used to filter Shared and Room Mailboxes.
.EXAMPLE
    .\M365_Resource_Report.ps1 -AdminUPN "admin@example.com" -DomainFilter "example.com"

.NOTES
    Requires PowerShell and the Exchange Online Management Module (e.g: Install-Module -Name ExchangeOnlineManagement)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$AdminUPN,

    [Parameter(Mandatory=$true)]
    [string]$DomainFilter
)

function Get-ResourceMailboxData {
    
    # --- 1. Connect to Exchange Online ---
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan

    try {
        # Connect-ExchangeOnline works natively on macOS/Core
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowProgress $false -ShowBanner:$false -ErrorAction Stop
        Write-Host "Connected to Exchange Online successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Exchange Online. Error: $($_.Exception.Message)"
        return
    }

    Write-Host "--- Starting Data Collection for domain: $DomainFilter ---" -ForegroundColor Yellow
    
    $ReportData = @()
    
    try {
        # FIX: Added 'RoomMailbox' to the RecipientTypeDetails list.
        Write-Host "Fetching Shared and Room Mailboxes..." -ForegroundColor Cyan
        
        $AllResources = Get-Mailbox -RecipientTypeDetails SharedMailbox,RoomMailbox -ResultSize Unlimited
        
        # Filter by the domain filter provided in parameters
        $TargetMailboxes = $AllResources | Where-Object { $_.PrimarySmtpAddress -like "*@$DomainFilter" }
        
        Write-Host "Found $($TargetMailboxes.Count) mailboxes (Shared + Room) matching domain '$DomainFilter'." -ForegroundColor Yellow
    }
    catch {
        Write-Error "Failed to fetch Mailboxes. Error: $($_.Exception.Message)"
        Disconnect-ExchangeOnline -Confirm:$false
        return
    }

    # --- 2. Process Each Mailbox ---
    $Counter = 0
    $Total = $TargetMailboxes.Count

    foreach ($Mbx in $TargetMailboxes) {
        $Counter++
        $PercentComplete = [math]::Round(($Counter / $Total) * 100)
        Write-Progress -Activity "Processing Mailboxes" -Status "Processing $($Mbx.PrimarySmtpAddress)" -PercentComplete $PercentComplete
        
        # --- A. Mailbox Size ---
        $MailboxSizeMB = 0
        try {
            # Get-MailboxStatistics retrieves the physical size
            $Stats = Get-MailboxStatistics -Identity $Mbx.Id -ErrorAction Stop | Select-Object TotalItemSize
            if ($Stats -and $Stats.TotalItemSize) {
                # Regex to extract raw bytes from format: "1.2 GB (1,234,567 bytes)"
                $SizeMatch = [regex]::Match($Stats.TotalItemSize.ToString(), '\(([\d,]+)\sbytes\)')
                $SizeInBytes = $SizeMatch.Groups[1].Value -replace ','
                if ($SizeInBytes -gt 0) {
                    $MailboxSizeMB = [math]::Round(([long]$SizeInBytes / 1MB), 2)
                }
            }
        }
        catch { 
            # If the mailbox has never been logged into, this might error or return null. default to 0.
            $MailboxSizeMB = 0 
        }

        # --- B. Aliases (Proxy Addresses) ---
        $AliasList = ""
        try {
            if ($Mbx.EmailAddresses) {
                # Filter for smtp addresses. 
                # 'SMTP:' (uppercase) is Primary. 'smtp:' (lowercase) is Alias.
                $Aliases = $Mbx.EmailAddresses | Where-Object { $_ -like "smtp:*" -and $_ -notlike "SMTP:*" }
                
                # Clean up the "smtp:" prefix
                $CleanAliases = $Aliases | ForEach-Object { $_ -replace "^smtp:", "" }
                $AliasList = $CleanAliases -join "; "
            }
        }
        catch { $AliasList = "Error" }

        # --- C. Create Custom Object ---
        $CustomObject = [PSCustomObject]@{
            DisplayName          = $Mbx.DisplayName
            PrimaryEmail         = $Mbx.PrimarySmtpAddress
            MailboxType          = $Mbx.RecipientTypeDetails # Identifies if it is Room or Shared
            MailboxSizeMB        = $MailboxSizeMB
            Aliases              = $AliasList
            LegacyExchangeDN     = $Mbx.LegacyExchangeDN
        }

        $ReportData += $CustomObject
    }

    # --- 3. Export ---
    Disconnect-ExchangeOnline -Confirm:$false

    $Date = Get-Date -Format "yyyyMMdd"
    $FileName = "ResourceMailbox_Report_$($DomainFilter)_$Date.csv"
    $ExportPath = Join-Path -Path $PWD -ChildPath $FileName

    Write-Host "Exporting data to $ExportPath..." -ForegroundColor Yellow
    $ReportData | Export-Csv -Path $ExportPath -NoTypeInformation
    Write-Host "Script complete." -ForegroundColor Green
}

# Run the function
Get-ResourceMailboxData -AdminUPN $AdminUPN -DomainFilter $DomainFilter