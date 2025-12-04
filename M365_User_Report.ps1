<# 
.SYNOPSIS
    Collects information on Users in a Microsoft 365 tenant filtered by a specified domain.
.DESCRIPTION
    This script connects to Microsoft Graph and Exchange Online, retrieves all users, filters them based on the provided domain, and collects details such as Display Name, Primary Email, Mailbox Size in MB, OneDrive Storage in MB, Group Memberships, Assigned Licenses, Aliases, and Last Successful Sign-In. The collected data is then exported to a CSV file for further analysis.
.PARAMETER TenantID
    The Tenant ID of the Microsoft 365 tenant.
.PARAMETER AdminUPN
    The User Principal Name (UPN) of the admin account used to connect to Exchange Online and Microsoft Graph.
.PARAMETER DomainFilter
    The domain used to filter users.
.EXAMPLE
    .\M365_User_Report.ps1 -TenantID "your-tenant-id" -AdminUPN "
.NOTES
    Requires PowerShell, the Microsoft Graph PowerShell SDK, and the Exchange Online Management Module.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$TenantID,

    [Parameter(Mandatory=$true)]
    [string]$AdminUPN,

    [Parameter(Mandatory=$true)]
    [string]$DomainFilter
)

function Get-M365UserMigrationData {
    # --- 1. Define License Mappings ---
    $LicenseMap = @{
        "STANDARDPACK"             = "Office 365 E1"
        "ENTERPRISEPACK"           = "Office 365 E3"
        "ENTERPRISEPREMIUM"        = "Office 365 E5"
        "SPB"                      = "Microsoft 365 Business Premium"
        "O365_BUSINESS_PREMIUM"    = "Microsoft 365 Business Standard"
        "O365_BUSINESS_ESSENTIALS" = "Microsoft 365 Business Basic"
        "ATP_ENTERPRISE"           = "Defender for Office 365 (Plan 1)"
        "FLOW_FREE"                = "Power Automate Free"
        "POWER_BI_STANDARD"        = "Power BI Free"
        "POWER_BI_PRO"             = "Power BI Pro"
        "TEAMS_EXPLORATORY"        = "Microsoft Teams Exploratory"
        "DESKLESSPACK"             = "Office 365 F3"
    }

    # --- 2. Connect ---
    Write-Host "Connecting to Microsoft Graph and Exchange Online..." -ForegroundColor Cyan

    try {
        Connect-MgGraph -TenantId $TenantID -Scopes "User.Read.All", "Group.Read.All", "Directory.Read.All", "Reports.Read.All", "AuditLog.Read.All", "Sites.Read.All", "Files.Read.All" -NoWelcome | Out-Null
        Write-Host "Connected to Microsoft Graph successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph. Error: $($_.Exception.Message)"
        return
    }
    
    try {
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowProgress $false -ShowBanner:$false
        Write-Host "Connected to Exchange Online successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Exchange Online. Error: $($_.Exception.Message)"
        Disconnect-MgGraph
        return
    }

    Write-Host "--- Starting Data Collection ---" -ForegroundColor Yellow
    
    $ReportData = @()
    
    try {
        # FIX 1: Explicitly request ProxyAddresses here. 
        # We exclude complex/dangerous properties but ensure we get the aliases.
        $Users = Get-MgUser -Filter "endsWith(userPrincipalName,'@$DomainFilter')" -All `
            -Property Id, UserPrincipalName, DisplayName, Mail, ProxyAddresses, AccountEnabled, SignInActivity `
            -ConsistencyLevel eventual
        
        Write-Host "Found $($Users.Count) users matching domain '$DomainFilter'." -ForegroundColor Yellow
    }
    catch {
        Write-Error "Failed to fetch users from Microsoft Graph. Error: $($_.Exception.Message)"
        Disconnect-ExchangeOnline
        Disconnect-MgGraph
        return
    }

    # --- 3. Process Each User ---
    $Counter = 0
    $Total = $Users.Count

    foreach ($User in $Users) {
        $Counter++
        $PercentComplete = [math]::Round(($Counter / $Total) * 100)
        Write-Progress -Activity "Processing Users" -Status "Processing $($User.UserPrincipalName)" -PercentComplete $PercentComplete
        
        # --- A. Mailbox Data ---
        $MailboxSizeMB = 0
        try {
            $MailboxData = Get-MailboxStatistics -Identity $User.UserPrincipalName -ErrorAction Stop | Select-Object TotalItemSize
            if ($MailboxData) {
                $SizeMatch = [regex]::Match($MailboxData.TotalItemSize.ToString(), '\(([\d,]+)\sbytes\)')
                $SizeInBytes = $SizeMatch.Groups[1].Value -replace ','
                if ($SizeInBytes -gt 0) {
                    $MailboxSizeMB = [math]::Round(([long]$SizeInBytes / 1MB), 2)
                }
            }
        }
        catch { $MailboxSizeMB = 0 }

        # --- B. OneDrive Storage (FIXED: Force Reload) ---
        $OneDriveUsage = 0
        try {
            # 1. Get the list of drives
            $AllDrives = Get-MgUserDrive -UserId $User.Id -All -ErrorAction SilentlyContinue
            # 2. Find the personal drive
            $PersonalDriveStub = $AllDrives | Where-Object { $_.DriveType -eq 'personal' } | Select-Object -First 1

            if ($PersonalDriveStub) {
                # 3. CRITICAL STEP: Re-fetch this specific drive to populate the Quota property
                $FullDrive = Get-MgDrive -DriveId $PersonalDriveStub.Id -ErrorAction SilentlyContinue
                
                if ($FullDrive -and $FullDrive.Quota) {
                    $UsedBytes = $FullDrive.Quota.Used
                    if ($UsedBytes -gt 0) {
                        $OneDriveUsage = [math]::Round(([long]$UsedBytes / 1MB), 2)
                    }
                }
            }
        }
        catch { }

        # --- C. Groups ---
        $GroupList = ""
        try {
            $UserGroups = Get-MgUserMemberOf -UserId $User.Id -All
            $GroupNames = @()
            foreach ($Group in $UserGroups) {
                if ($Group.DisplayName) { $GroupNames += $Group.DisplayName }
                elseif ($Group.AdditionalProperties['displayName']) { $GroupNames += $Group.AdditionalProperties['displayName'] }
            }
            $GroupList = $GroupNames -join "; "
        }
        catch { $GroupList = "Error" }

        # --- D. Licenses (Translated) ---
        $ReadableLicenseList = ""
        try {
            $Licenses = Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction SilentlyContinue
            if ($Licenses) {
                $TranslatedNames = @()
                foreach ($Lic in $Licenses) {
                    $Sku = $Lic.SkuPartNumber
                    if ($LicenseMap.ContainsKey($Sku)) { $TranslatedNames += $LicenseMap[$Sku] }
                    else { $TranslatedNames += $Sku }
                }
                $ReadableLicenseList = $TranslatedNames -join "; "
            }
        }
        catch { $ReadableLicenseList = "Error" }

        # --- E. Aliases (FIXED: Simplified Logic) ---
        $AliasList = ""
        # Now that we requested ProxyAddresses in the main query, this should populate.
        if ($User.ProxyAddresses) {
            # Filter for lowercase 'smtp' which usually denotes aliases (Uppercase SMTP is primary)
            $Aliases = $User.ProxyAddresses | Where-Object { $_ -match "^smtp:" }
            
            # If you want ALL addresses (Primary + Aliases), remove the '-match' check above.
            # Clean up the "smtp:" prefix for cleaner CSV output
            $CleanAliases = $Aliases | ForEach-Object { $_ -replace "^smtp:", "" -replace "^SMTP:", "" }
            $AliasList = $CleanAliases -join "; "
        }

        # --- F. Last Sign In ---
        $LastSignIn = $null
        if ($User.SignInActivity) {
            $LastSignIn = $User.SignInActivity.LastSuccessfulSignInDateTime
        }
        elseif ($User.AdditionalProperties['signInActivity']) {
            $LastSignIn = $User.AdditionalProperties['signInActivity']['lastSuccessfulSignInDateTime']
        }

        # --- G. Create Custom Object ---
        $CustomObject = [PSCustomObject]@{
            SourceEmail                  = $User.UserPrincipalName
            DisplayName                  = $User.DisplayName
            AccountEnabledAtSource       = $User.AccountEnabled
            LastSuccessfulSignIn         = $LastSignIn
            MailboxSizeMB                = $MailboxSizeMB
            OneDriveStorageMB            = $OneDriveUsage
            IsGuestUser                  = ($User.UserType -eq "Guest")
            AssignedLicensesAtSource     = $ReadableLicenseList
            GroupMembershipsAtSource     = $GroupList
            SourceAddresses              = $AliasList
        }

        $ReportData += $CustomObject
    }

    # --- 4. Export ---
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph

    $Date = Get-Date -Format "yyyyMMdd"
    $FileName = "M365_User_Report_$($DomainFilter)_$Date.csv"
    $ExportPath = Join-Path -Path $PWD -ChildPath $FileName

    Write-Host "Exporting data to $ExportPath..." -ForegroundColor Yellow
    $ReportData | Export-Csv -Path $ExportPath -NoTypeInformation
    Write-Host "Script complete." -ForegroundColor Green
}

Get-M365UserMigrationData -TenantID $TenantID -AdminUPN $AdminUPN -DomainFilter $DomainFilter