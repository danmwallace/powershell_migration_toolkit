<#
.SYNOPSIS
    Collects information on Microsoft 365 Groups in a tenant filtered by a specified domain.
.DESCRIPTION
    This script connects to Exchange Online, retrieves all Microsoft 365 Groups (Unified Groups), filters
    them based on the provided domain, and collects details such as Display Name, Primary Email, Mailbox Size in MB,
    SharePoint Site URL, Owner and Member counts, and whether the group is connected to a
    Microsoft Team. The collected data is then exported to a CSV file for further analysis.
.PARAMETER TenantID
    The Tenant ID of the Microsoft 365 tenant.
.PARAMETER AdminUPN
    The User Principal Name (UPN) of the admin account used to connect to Exchange Online.
.PARAMETER DomainFilter
    The domain used to filter Microsoft 365 Groups.
.EXAMPLE
    .\M365_Group_Report.ps1 -TenantID "your-tenant-id" -AdminUPN "admin@yourdomain.com" -DomainFilter "yourdomain.com"
.NOTES
    Requires PowerShell (Cross-platform: Windows, macOS, Linux) and the Exchange Online Management Module (Install-Module -Name ExchangeOnlineManagement).
#>


[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$TenantID, # Kept for parameter consistency, though not strictly needed for EXO connection

    [Parameter(Mandatory=$true)]
    [string]$AdminUPN,

    [Parameter(Mandatory=$true)]
    [string]$DomainFilter
)

function Get-M365GroupMigrationData {
    
    # --- 1. Connect to Exchange Online ---
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan

    try {
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowProgress $false -ShowBanner:$false -ErrorAction Stop
        Write-Host "Connected to Exchange Online successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Exchange Online. Error: $($_.Exception.Message)"
        return
    }

    Write-Host "--- Starting M365 Group Data Collection for domain: $DomainFilter ---" -ForegroundColor Yellow
    
    $ReportData = @()
    
    try {
        Write-Host "Fetching all Microsoft 365 Groups (Unified Groups)..." -ForegroundColor Cyan
        
        # We fetch all groups first, then filter. 
        # Server-side filtering on specific string matches can sometimes be tricky with UnifiedGroups, 
        # but filtering locally ensures accuracy for the DomainFilter.
        $AllGroups = Get-UnifiedGroup -ResultSize Unlimited
        
        $TargetGroups = $AllGroups | Where-Object { $_.PrimarySmtpAddress -like "*@$DomainFilter" }
        
        Write-Host "Found $($TargetGroups.Count) M365 Groups matching domain '$DomainFilter'." -ForegroundColor Yellow
    }
    catch {
        Write-Error "Failed to fetch Groups. Error: $($_.Exception.Message)"
        Disconnect-ExchangeOnline -Confirm:$false
        return
    }

    # --- 2. Process Each Group ---
    $Counter = 0
    $Total = $TargetGroups.Count

    foreach ($Group in $TargetGroups) {
        $Counter++
        $PercentComplete = [math]::Round(($Counter / $Total) * 100)
        Write-Progress -Activity "Processing M365 Groups" -Status "Processing $($Group.DisplayName)" -PercentComplete $PercentComplete
        
        # --- A. Group Mailbox Size ---
        $MailboxSizeMB = 0
        try {
            $Stats = Get-MailboxStatistics -Identity $Group.Id -ErrorAction Stop | Select-Object TotalItemSize
            if ($Stats -and $Stats.TotalItemSize) {
                # Regex to extract raw bytes
                $SizeMatch = [regex]::Match($Stats.TotalItemSize.ToString(), '\(([\d,]+)\sbytes\)')
                $SizeInBytes = $SizeMatch.Groups[1].Value -replace ','
                if ($SizeInBytes -gt 0) {
                    $MailboxSizeMB = [math]::Round(([long]$SizeInBytes / 1MB), 2)
                }
            }
        }
        catch { $MailboxSizeMB = 0 }

        # --- B. Owners & Members ---
        $OwnerList = ""
        $MemberList = ""
        $MemberCount = 0

        try {
            # Get Owners
            $GroupOwners = Get-UnifiedGroupLinks -Identity $Group.Id -LinkType Owners -ResultSize Unlimited -ErrorAction SilentlyContinue
            if ($GroupOwners) {
                $OwnerList = ($GroupOwners.PrimarySmtpAddress) -join "; "
            }

            # Get Members
            $GroupMembers = Get-UnifiedGroupLinks -Identity $Group.Id -LinkType Members -ResultSize Unlimited -ErrorAction SilentlyContinue
            if ($GroupMembers) {
                $MemberCount = $GroupMembers.Count
                # We limit the string length for CSV readability. If huge, we might truncate or just list count.
                # Here we list email addresses separated by semicolons.
                $MemberList = ($GroupMembers.PrimarySmtpAddress) -join "; "
            }
        }
        catch { 
            $OwnerList = "Error fetching links" 
        }

        # --- C. Teams & SharePoint Status ---
        $IsTeam = $false
        if ($Group.ResourceProvisioningOptions -contains "Team") {
            $IsTeam = $true
        }

        # --- D. Create Custom Object ---
        $CustomObject = [PSCustomObject]@{
            GroupName            = $Group.DisplayName
            PrimaryEmail         = $Group.PrimarySmtpAddress
            IsMicrosoftTeam      = $IsTeam
            MailboxSizeMB        = $MailboxSizeMB
            SharePointSiteUrl    = $Group.SharePointSiteUrl
            OwnerCount           = $GroupOwners.Count
            MemberCount          = $MemberCount
            OwnerEmails          = $OwnerList
            MemberEmails         = $MemberList
            AccessType           = $Group.AccessType # Private or Public
        }

        $ReportData += $CustomObject
    }

    # --- 3. Export ---
    Disconnect-ExchangeOnline -Confirm:$false

    $Date = Get-Date -Format "yyyyMMdd"
    $FileName = "M365_Group_Report_$($DomainFilter)_$Date.csv"
    $ExportPath = Join-Path -Path $PWD -ChildPath $FileName

    Write-Host "Exporting data to $ExportPath..." -ForegroundColor Yellow
    $ReportData | Export-Csv -Path $ExportPath -NoTypeInformation
    Write-Host "Script complete." -ForegroundColor Green
}

# Run the function
Get-M365GroupMigrationData -TenantID $TenantID -AdminUPN $AdminUPN -DomainFilter $DomainFilter