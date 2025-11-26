<#
.SYNOPSIS
    Updates the Source or Destination tenant by modifying user email addresses and account statuses based on a CSV file (UsersCsv). 
    Provide the Tenant configuration (Tenant IDs, Admin UPNs) via a second CSV file (TenantsCsv).
    Is designed to configure identities based on the values provided in the CSV file. See the readme for more details.
.DESCRIPTION
    This script connects to Exchange Online and converts emails for users based on a provided CSV file. It updates the PrimarySMTP address and removes specified aliases.
.PARAMETER UsersCsv
    The path to the CSV file containing the user email conversion data. Spreadsheet must include SourceEmail, PostMigrationSourceEmail, DestinationAlias, DestinationStagingEmail, PostMigrationDestinationEmail, DestinationPassword, AccountEnabledAtSource, and AccountEnabledAtDestination columns.
.PARAMETER TenantsCsv
    The path to the CSV file containing the tenant configuration data. Spreadsheet must include SourceTenantID, SourceTenantAdmin, DestinationTenantID, and DestinationTenantAdmin columns.
.PARAMETER Target
    Specify whether we are updating the 'Source' or 'Destination' tenant.
.PARAMETER Revert
    (Optional) If specified, reverts the changes made during migration by using pre-migration email addresses from the CSV file. Does not revert $AccountStatus or remove aliases that may have been added previously. If you need to revert $AccountStatus, you will want to update the $true/$false values in the CSV file accordingly.
.EXAMPLE
    .\UpdateMigratedAccounts.ps1 -UsersCsv "C:\Path\To\Users.csv" -TenantsCsv "C:\Path\To\Tenants.csv" -Target "Source"
    Updates the Source tenant based on the user email mappings provided in Users.csv and tenant configuration in Tenants.csv.
.EXAMPLE
    .\UpdateMigratedAccounts.ps1 -UsersCsv "C:\Path\To\Users.csv" -TenantsCsv "C:\Path\To\Tenants.csv" -Target "Destination"
    Updates the Destination tenant based on the user email mappings provided in Users.csv and tenant configuration in Tenants.csv.
.EXAMPLE
    .\UpdateMigratedAccounts.ps1 -UsersCsv "C:\Path\To\Users.csv" -TenantsCsv "C:\Path\To\Tenants.csv" -Target "Source" -Revert
    Reverts the email address changes in the Source tenant based on the pre-migration email addresses provided in Users.csv.
.EXAMPLE
    .\UpdateMigratedAccounts.ps1 -UsersCsv "C:\Path\To\Users.csv" -TenantsCsv "C:\Path\To\Tenants.csv" -Target "Destination" -Revert
    Reverts the email address changes in the Destination tenant based on the pre-migration email addresses provided in Users.csv.
.NOTES
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$UsersCsv,

    [Parameter(Mandatory=$true)]
    [string]$TenantsCsv,

    [Parameter(Mandatory=$true)]
    [string]$Target,

    [Parameter(Mandatory=$false)]
    [switch]$Revert
)

$FormerAddresses = "FormerAddresses_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$FormerAddressResults = @()

# Check if the User mapping file exists before proceeding
if (-not (Test-Path $UsersCsv)) {
    Write-Error "User mapping CSV file not found at: $UsersCsv"
    Write-Error "Please create a CSV file with the following headers: SourceEmail,PostMigrationSourceEmail,DestinationAlias,DestinationStagingEmail,PostMigrationDestinationEmail,DestinationPassword,AccountEnabledAtSource, and AccountEnabledAtDestination values"
    Write-Error "See the readme for more details."
    exit
}

# Check if the Tenant configuration file exists before proceeding
if (-not (Test-Path $TenantsCsv)) {
    Write-Error "Tenant configuration CSV file not found at: $TenantsCsv"
    Write-Error "Please create a CSV file with the following headers: SourceTenantID,SourceTenantAdmin,DestinationTenantID,DestinationTenantAdmin"
    Write-Error "See the readme for more details."
    exit
}

# 3. Import the CSV content
# We use a variable name that clearly indicates the collection of data
$MailboxData = Import-Csv -Path $UsersCsv
$TenantData = Import-Csv -Path $TenantsCsv

# Identify if we're connecting to the Source Tenant or the Destination Tenant and set $TenantID and $UserPrincipalName accordingly
if ($Target -eq "Source") {
    $TenantID = $TenantData.SourceTenantID
    $UserPrincipalName = $TenantData.SourceTenantAdmin
}
elseif ($Target -eq "Destination") {
    $TenantID = $TenantData.DestinationTenantID
    $UserPrincipalName = $TenantData.DestinationTenantAdmin
}
else {
    Write-Error "Invalid Target specified. Please specify whether we are updating the 'Source' or 'Destination' when running the script."
    exit
}

# Connect with MG Graph and ExchangeOnline(Management) before running the script
Connect-MGGraph -Scopes "User.ReadWrite.All" -TenantId $TenantID -NoWelcome
Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName

# Iterate through each mailbox entry in the CSV
foreach ($Mailbox in $MailboxData) {
    # Determine if the tenant we are modifying is the Source or Destination Tenant, and set variables accordingly   
    if ($Target -eq "Source") {
        # If -Revert is specified, swap the email addresses to revert previous changes
        if ($Revert) {
            $Identity = $Mailbox.PostMigrationSourceEmail
            $Email = $Mailbox.SourceEmail
        }
        # Otherwise, use the normal source emails. The $Identity is the current email, and $Email is the intended email post-migration
        else {
            $Identity = $Mailbox.SourceEmail
            $Email = $Mailbox.PostMigrationSourceEmail
        }
        $AccountStatusString = $Mailbox.AccountEnabledAtSource
    }
    elseif ($Target -eq "Destination") {
        # If -Revert is specified, swap the email addresses to revert previous changes
        if ($Revert) {
            $Identity = $Mailbox.PostMigrationDestinationEmail
            $Email = $Mailbox.DestinationStagingEmail
        }
        # Otherwise, use the normal destination emails. The $Identity is the current email, and $Email is the intended email post-migration
        else {
            $Identity = $Mailbox.DestinationStagingEmail
            $Email = $Mailbox.PostMigrationDestinationEmail
            $Alias = $Mailbox.DestinationAlias
            $TemporaryPassword = $Mailbox.DestinationPassword
        }
        $AccountStatusString = $Mailbox.AccountEnabledAtDestination
    }
    try {
        # Retrieve the EmailAddresses property
        $ExistingAddresses = (Get-Mailbox -Identity $Identity -ErrorAction Stop).EmailAddresses
        
        # Filter for Aliases: Select addresses that DO NOT start with 'SMTP:' (Primary) 
        # and DO NOT start with 'sip:' (Skype/Teams addresses).
        $Aliases = $ExistingAddresses | Where-Object { 
            $_ -notlike "SMTP:*" -and $_ -notlike "sip:*"
        } | ForEach-Object { 
            # Convert the complex object back to a simple string (e.g., 'X400:c=us...')
            $_.AddressString 
        }
        
        # Join the array of alias strings into a single comma-separated string for the CSV cell
        $AliasString = $Aliases -join ", "
        
        # Create a Custom Object for Export
        $CustomObject = [PSCustomObject]@{
            Identity = $Identity
            Aliases  = $AliasString
        }
        
        # Add the custom object to the results array
        $FormerAddressResults += $CustomObject
    }
    catch {
        Write-Warning "Failed to process identity '$Identity': $($_.Exception.Message)"
        # Create an object for failed users as well
        $FormerAddressResults += [PSCustomObject]@{
            Identity = $Identity
            Aliases  = "ERROR: $($_.Exception.Message)"
        }
    }
    # Convert the Account Status string to a boolean by using an explicit comparison. This is required for it to work later when passed via Update-MgUser
    $AccountStatus = ($AccountStatusString -eq "True")
    $AccountParams = @{
        accountEnabled = $AccountStatus
    }
    # Password Profile (only used for the Destination tenant). We also set it so the user must change their password on next sign-in. We don't control this in the CSV as it is a pretty common expectation.
    $PasswordProfile = @{
        Password                      = $TemporaryPassword
        ForceChangePasswordNextSignIn = $true
    }
    # Inside the foreach loop, after setting $Identity and $Email:
    try {
        # Get the user object using the current identifier ($Identity, which is the current email/UPN)
        $UserObject = Get-MgUser -UserId $Identity -Property UserPrincipalName, Id
        $CurrentUPN = $UserObject.UserPrincipalName
        $ObjectID = $UserObject.Id
    }
    catch {
        Write-Error "FAILURE: Could not retrieve user object for '$Identity'. Error: $($_.Exception.Message)"
        continue
    }
    
    Write-Host "Processing User: $Identity ..." -ForegroundColor Yellow
    if ([string]::IsNullOrEmpty($Identity) -or [string]::IsNullOrEmpty($Email)) {
        Write-Warning "Skipping row: Identity or Email is missing in a row. Please check the CSV file."
        continue
    }
    else {
        # Change Account Status based on $AccountStatus in spreadsheet
        try {
            Write-Host "Changing account enabled status for '$Identity' to '$AccountStatus'" -ForegroundColor White
            Update-MgUser -UserId $ObjectID -BodyParameter $AccountParams -ErrorAction Stop
            Write-Host "SUCCESS: Account status for '$Identity' has been updated to '$AccountStatus'" -ForegroundColor Green
        }
        catch {
            Write-Error "FAILURE: Could not modify account status for '$SourceEmail'. Error: $($_.Exception.Message)"
            continue
        }
        try { 
            Write-Host "Setting password for '$Identity'" -ForegroundColor White
            Update-MgUser -UserId $ObjectID -PasswordProfile $PasswordProfile -ErrorAction Stop
            Write-Host "SUCCESS: Set password for '$Identity' to '$TemporaryPassword'" -ForegroundColor Green
        }
        catch { 
            Write-Error "FAILURE: Could not set password for '$Identity'. Error: $($_.Exception.Message)"
            continue
        }
        # Change the user identity (UPN) based on the spreadsheet
        try { 
            Write-Host "Changing User Principal Name (UPN) for '$Identity' to '$Email'" -ForegroundColor White
            Update-MgUser -UserId $ObjectID -UserPrincipalName $Email -Mail $Email -ErrorAction Stop
            Write-Host "SUCCESS: Changed UPN for '$Identity' to '$Email'" -ForegroundColor Green
        }
        catch { 
            Write-Error "FAILURE: Could not change Email for '$Identity' to '$Email'. Error: $($_.Exception.Message)"
            continue
        }
        # Change the actual email address, e.g $PrimarySMTP and $WindowsEmailAddress values in Exchange. This is also where the $Alias value is added if provided.
        try {
            Write-Host "Changing email address from $Identity to $Email" -ForegroundColor White
            Set-Mailbox -Identity $Identity -EmailAddresses $NewAddresses -WindowsEmailAddress $Email -MicrosoftOnlineServicesID $Email -ErrorAction Stop
        } 
        catch {
            Write-Error "FAILURE: Could not change email address from '$Identity' to '$Email'. Error: $($_.Exception.Message)"
            continue
        }
        # Remove the old address from the mailbox.
        #try {
        #    Write-Host "Removing old email address '$Identity' from mailbox." -ForegroundColor White
        #    Set-Mailbox -Identity $Identity -EmailAddresses @{remove = "$Identity" } -ErrorAction Stop
        #}
        #catch {
        #    Write-Error "FAILURE: Could not remove old email address '$Identity' from mailbox. Error: $($_.Exception.Message)"
        #    continue
        #}
    }
    $FormerAddressResults | Export-Csv -Path "$FormerAddresses.csv" -NoTypeInformation
}

# Disconnect from Exchange Online and MG Graph and end the script
#Disconnect-ExchangeOnline -Confirm:$false
#Disconnect-MgGraph
Write-Host "`nScript execution complete."