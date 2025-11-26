# Requires:
# 1. PowerShell installed on macOS
# 2. Microsoft Graph PowerShell SDK: Install-Module Microsoft.Graph -Scope CurrentUser
# 3. Exchange Online Management: Install-Module ExchangeOnlineManagement -Scope CurrentUser

param(
    [Parameter(Mandatory=$true)]
    [string]$TenantID,

    [Parameter(Mandatory=$true)]
    [string]$AdminUPN,

    [Parameter(Mandatory=$true)]
    [string]$DomainFilter
)

function Get-M365UserMigrationData {
    # --- 1. Connect to Microsoft Graph and Exchange Online ---
    Write-Host "Connecting to Microsoft Graph and Exchange Online..." -ForegroundColor Cyan

    # **IMPORTANT:** Added 'AuditLog.Read.All' scope for Last Sign In data (SignInActivity)
    try {
        Connect-MgGraph -TenantId $TenantID -Scopes "User.Read.All", "Group.Read.All", "Directory.Read.All", "Reports.Read.All", "AuditLog.Read.All" -Account $AdminUPN
        Write-Host "Connected to Microsoft Graph successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph. Check TenantID, AdminUPN, and permissions. Error: $($_.Exception.Message)"
        return
    }

    # Connect-ExchangeOnline for Mailbox details (size, aliases)
    try {
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowProgress $false
        Write-Host "Connected to Exchange Online successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Exchange Online. Check AdminUPN, and make sure the EXO V3 module is installed. Error: $($_.Exception.Message)"
        Disconnect-MgGraph
        return
    }

    Write-Host "--- Starting Data Collection ---" -ForegroundColor Yellow
    
    $ReportData = @()

    # --- 2. Get Users based on Domain Filter ---
    Select-MgProfile -Name "v1.0"
    
    try {
        # **MODIFIED:** Added 'AccountEnabled' and 'SignInActivity' to the properties list
        $Users = Get-MgUser -Filter "endsWith(userPrincipalName,'@$DomainFilter')" -All `
            -Property Id, UserPrincipalName, DisplayName, Mail, LicenseDetails, ProxyAddresses, UserType, AccountEnabled, SignInActivity
        
        Write-Host "Found $($Users.Count) users matching domain '$DomainFilter'." -ForegroundColor Yellow
    }
    catch {
        Write-Error "Failed to fetch users from Microsoft Graph. Error: $($_.Exception.Message)"
        Disconnect-ExchangeOnline
        Disconnect-MgGraph
        return
    }

    # --- 3. Process Each User ---
    foreach ($User in $Users) {
        Write-Host "Processing $($User.UserPrincipalName)..."
        
        # --- A. Mailbox Data (from Exchange Online) ---
        $MailboxData = Get-MailboxStatistics -Identity $User.UserPrincipalName -ErrorAction SilentlyContinue | Select-Object TotalItemSize, PrimarySmtpAddress

        $MailboxSizeMB = 0
        if ($MailboxData) {
            $SizeInBytes = [regex]::Match($MailboxData.TotalItemSize.ToString(), '\(([\d,]+)\sbytes\)').Groups[1].Value -replace ','
            if ($SizeInBytes -and ($SizeInBytes -ne 0)) {
                $MailboxSizeMB = [math]::Round(($SizeInBytes / 1MB), 2)
            }
        }

        # --- B. OneDrive Storage (from Microsoft Graph Reports) ---
        $OneDriveUsage = 0 # Default to 0 MB
        try {
            $Report = Get-MgReportOneDriveUsageFile -Period 'D7' -OutFile $null -ErrorAction Stop
            $ReportObj = $Report | ConvertFrom-Csv
            
            $UserUsageRecord = $ReportObj | Where-Object { $_.UserPrincipalName -eq $User.UserPrincipalName } | Sort-Object Date -Descending | Select-Object -First 1
            
            if ($UserUsageRecord -and $UserUsageRecord.StorageUsed -ne 0) {
                $OneDriveUsage = [math]::Round(($UserUsageRecord.StorageUsed / 1MB), 2)
            }
        }
        catch {
            Write-Warning "Could not retrieve OneDrive usage for $($User.UserPrincipalName). (Requires Reports Reader role)"
        }

        # --- C. Groups, Aliases, and Licenses (from Microsoft Graph) ---
        
        $UserGroups = Get-MgUserMemberOf -UserId $User.Id -All | Select-Object -ExpandProperty DisplayName
        $GroupList = $UserGroups -join "; "

        $LicenseList = $User.LicenseDetails.SkuPartNumber -join "; "

        $Aliases = @($User.ProxyAddresses) | Where-Object { $_ -match "^smtp:" -and $_ -notmatch "^SMTP:" }
        $AliasList = ($Aliases -replace "^smtp:", "") -join "; "

        # --- D. NEW: Last Sign In and Account Status ---
        
        # LastSuccessfulSignInDateTime is generally preferred as it indicates real user activity.
        # It will be $null if the user has never signed in.
        $LastSignIn = $null
        if ($User.SignInActivity) {
            $LastSignIn = $User.SignInActivity.LastSuccessfulSignInDateTime
        }


        # --- E. Create Custom Object and Add to Array ---
        $CustomObject = [PSCustomObject]@{
            UPN                     = $User.UserPrincipalName
            DisplayName             = $User.DisplayName
            AccountEnabled          = $User.AccountEnabled # NEW FIELD
            LastSuccessfulSignIn    = $LastSignIn          # NEW FIELD
            MailboxSizeMB           = $MailboxSizeMB
            OneDriveStorageMB       = $OneDriveUsage
            IsGuestUser             = ($User.UserType -eq "Guest")
            AssignedLicenses        = $LicenseList
            GroupMemberships        = $GroupList
            EmailAliases            = $AliasList
        }

        $ReportData += $CustomObject
    }

    # --- 4. Disconnect and Export ---
    Write-Host "--- Data Collection Complete ---" -ForegroundColor Green
    
    Disconnect-ExchangeOnline
    Disconnect-MgGraph

    # Export to CSV
    $Date = Get-Date -Format "yyyyMMdd"
    $FileName = "M365_Migration_Report_$($DomainFilter)_$Date.csv"

    Write-Host "Exporting data to $FileName..." -ForegroundColor Yellow
    $ReportData | Export-Csv -Path $FileName -NoTypeInformation

    Write-Host "Script complete. Output file created: $FileName" -ForegroundColor Green
}

Get-M365UserMigrationData -TenantID $TenantID -AdminUPN $AdminUPN -DomainFilter $DomainFilter