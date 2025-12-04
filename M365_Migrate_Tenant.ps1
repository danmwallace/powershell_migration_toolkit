<# PARAMS
.SYNOPSIS
    Migrate user accounts between Microsoft 365 tenants by updating email addresses, account status, and other properties based on CSV input.
.DESCRIPTION
    This script reads user and tenant configuration from CSV files and updates user accounts in the specified Microsoft
    365 tenant (Source or Destination) according to the provided mappings. It supports dry-run mode for testing changes
    without applying them, and can revert changes if needed.
.PARAMETER UsersCsv
    Path to the CSV file containing user mappings and properties.
.PARAMETER TenantsCsv
    Path to the CSV file containing tenant configuration details.
.PARAMETER Target
    Specifies whether to update the 'Source' or 'Destination' tenant. Source being the tenant where users are migrating from, and Destination being where they are migrating to. Configure $TenantsCsv accordingly.
.PARAMETER TargetDomain
    The domain being migrated. This is used to identify and remove email addresses associated with this domain when updating user accounts in the Source Tenant so you can remove the domain in Microsoft 365 later.
.PARAMETER Revert
    Switch to indicate if the script should revert previous changes. This will swap email addresses back to their original values. Will not re-add removed aliases that are associated with the $TargetDomain.
.PARAMETER Dryrun
    Switch to enable dry-run mode, where no changes are made but actions are logged.
.EXAMPLE
    .\MigrateUsers.ps1 -UsersCsv "C:\path\to\users.csv" -TenantsCsv "C:\path\to\tenants.csv" -Target "Source" -TargetDomain "example.com" -Dryrun
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$UsersCsv,

    [Parameter(Mandatory=$true)]
    [string]$TenantsCsv,

    [Parameter(Mandatory=$true)]
    [string]$Target,

    [Parameter(Mandatory=$true)]
    [string]$TargetDomain,

    [Parameter(Mandatory=$false)]
    [switch]$Revert,

    [Parameter(Mandatory=$false)]
    [switch]$Dryrun
)

$FormerAddresses = "FormerAddresses_$($TargetDomain)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$FormerAddressResults = @()

############################# [PRE-FLIGHT CHECKS AND IMPORTS] #############################
# Check if required modules are installed
# If not, inform the user to install them

$RequiredModules = @("Microsoft.Graph.Users", "ExchangeOnlineManagement")
foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        Write-Error "Required module '$Module' is not installed. Please install it using 'Install-Module $Module -Scope CurrentUser' and try again."
        exit
    }
}

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

############################## [IMPORT CSV FILES AND SETUP CONNECTIONS] #############################
# Import the CSV content
# We use a variable name that clearly indicates the collection of data

$MailboxData = Import-Csv -Path $UsersCsv
$TenantData = Import-Csv -Path $TenantsCsv

# Identify if we're connecting to the Source Tenant or the Destination Tenant and set $TenantID and $UserPrincipalName accordingly
if ($Target -eq "Source") {
    $TenantID = $TenantData.SourceTenantID
    $UserPrincipalName = $TenantData.SourceAdminAccount
}
elseif ($Target -eq "Destination") {
    $TenantID = $TenantData.DestinationTenantID
    $UserPrincipalName = $TenantData.DestinationAdminAccount
}
else {
    Write-Error "Invalid Target specified. Please specify whether we are updating the 'Source' or 'Destination' when running the script."
    exit
}

############################# [CONNECT TO TENANT] #############################
# Inform the user about the mode they are running the script in, and the tenant they are connecting to
# If -Dryrun is not specified, we are in live migration mode, and the User is made to delay for 5 seconds before proceeding

if (-not $Dryrun) {
    Write-Host "================= MIGRATION MODE ENABLED ================" -ForegroundColor Magenta
    Write-Host "Making changes on '$Target' tenant..." -ForegroundColor Magenta
    Write-Host "Tenant ID: $TenantID" -ForegroundColor Magenta
    Write-Host "Admin: $UserPrincipalName" -ForegroundColor Magenta
    Write-Host " "
    Write-Host "Please sign in with your Admin credentials when prompted." -ForegroundColor Magenta
    Write-Host " "
}
else {
    Write-Host "================= DRY RUN MODE ENABLED ================" -ForegroundColor Cyan
    Write-Host "Making changes on '$Target' tenant..." -ForegroundColor Cyan
    Write-Host "Tenant ID: $TenantID" -ForegroundColor Cyan
    Write-Host "Admin: $UserPrincipalName" -ForegroundColor Cyan
    Write-Host " "
    Write-Host "Please sign in with your Admin credentials when prompted." -ForegroundColor Cyan
    Write-Host " "
}

# Connect with MG Graph and ExchangeOnline(Management) before running the script
Connect-MGGraph -Scopes "User.ReadWrite.All" -TenantId $TenantID -NoWelcome | Out-Null
Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false

####################### [ DRY RUN CHECK AND VERIFICATION ] #######################
if (-not $Dryrun) {
    Write-Host "WARNING: You are about to make LIVE changes to the '$Target' tenant!" -ForegroundColor Red
    Write-Host "If this was a mistake, you can stop the script now (Ctrl+C) or press Enter to continue..." -ForegroundColor Magenta
    Write-Host "Please perform a -dryrun before running live changes if you have not already done so!" -ForegroundColor Magenta
    Write-Host "You have been warned! Waiting 10 seconds, press CTRL+C/CMD+C to cancel..." -ForegroundColor Red
    Write-Host " "
    Start-Sleep -Seconds 10
    Write-Host "Press enter to continue..." -ForegroundColor Red
    Read-Host
}
else {
    Write-Host "Please observe changes in the log output and adjust your CSV as necessary" -ForegroundColor Cyan
    Write-Host "If you would like to cancel, you can stop the script now with CTRL+C (Windows) or CMD+C (macOS)..." -ForegroundColor Cyan
    Write-Host " "
    Write-Host "Press enter to continue..." -ForegroundColor Cyan
    Read-Host
}

############################# [SOURCE AND DESTINATION MAPPINGS] #############################
# Iterate through each mailbox entry in the CSV
# This is where the logic for -Target and -Revert is applied
# We set the $Identity and $Email variables based on the Target and Revert parameters
# If -Revert is specified, we swap the email addresses to revert (presumably) previous changes

foreach ($Mailbox in $MailboxData) {
    # Determine if the tenant we are modifying is the Source or Destination Tenant, and set variables accordingly   
    if ($Target -eq "Source") {
        $AccountHiddenFromGALString = $Mailbox.SourceHideFromGAL
        $AccountStatusString = $Mailbox.AccountEnabledAtSource
        $MailboxType = $Mailbox.SourceMailboxType
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
    }
    elseif ($Target -eq "Destination") {
        # If -Revert is specified, swap the email addresses to revert previous changes
        $AccountHiddenFromGALString = $Mailbox.DestinationHideFromGAL
        $TemporaryPassword = $Mailbox.DestinationPassword
        $AccountStatusString = $Mailbox.AccountEnabledAtDestination
        $MailboxType = $Mailbox.DestinationMailboxType
        if ($Revert) {
            $Identity = $Mailbox.PostMigrationDestinationEmail
            $Email = $Mailbox.DestinationStagingEmail
        }
        # Otherwise, use the normal destination emails. The $Identity is the current email, and $Email is the intended email post-migration
        else {
            $Identity = $Mailbox.DestinationStagingEmail
            $Email = $Mailbox.PostMigrationDestinationEmail
            $Aliases = $Mailbox.DestinationAliases
        }
    }


    ############################### [PROXY ADDRESS PROCESSING BLOCK] ##########################
    try {
        [System.Collections.Generic.List[string]]$AddressesToKeepList = @()
        Write-Host "================= COMPILING PROXY ADDRESSES ================" -ForegroundColor DarkYellow
        Write-Host "Compiling proxy addresses for '$Identity' ..." -ForegroundColor DarkYellow
        
        # Retrieve the EmailAddresses property to get existing proxy addresses
        $ExistingAddresses = (Get-Mailbox -Identity $Identity -ErrorAction Stop).EmailAddresses
        
        # Join the collection directly to display all addresses in one line
        Write-Host "  -> Existing Addresses for '$Identity' include: $($ExistingAddresses -join '; ')" -ForegroundColor DarkYellow
        
        # Create the Custom Object for Exporting to a CSV report
        $FormerProxyAddresses = [PSCustomObject]@{
            Identity = $Identity
            Addresses = $ExistingAddresses -join '; ' # Join for cleaner CSV output
        }
        # Add the custom object to the results array for exporting the existing aliases
        $FormerAddressResults += $FormerProxyAddresses

        # Filter out any existing addresses that match the target domain. We want to remove those belonging to the domain, but keep others
        foreach ($AddressObject in $ExistingAddresses) {
            $AddressString = $AddressObject.ToString()
            # If the address is NOT the target domain AND it is NOT the new primary email (to prevent duplicates), KEEP it
            if (($AddressString -notlike "*@$TargetDomain") -and ($AddressString -ne "smtp:$Email")) {
                Write-Host "  -> Keeping address: $AddressString" -ForegroundColor DarkYellow
                # FIX: Use .Add() method on the List
                $AddressesToKeepList.Add($AddressString)
            }
            else {
                Write-Host "  -> Discarding address: $AddressString" -ForegroundColor DarkYellow
            }
        }

        # Add aliases/proxy addresses from the CSV (if present)
        if ($Aliases -ne "" -and $null -ne $Aliases) {
            # Split the semicolon-separated string into an array, clean up whitespace
            $NewAliases = $Aliases.Split(';') | ForEach-Object { $_.Trim() }
            
            Write-Host "  -> Adding $($NewAliases.Count) secondary aliases..." -ForegroundColor DarkYellow
            foreach ($Alias in $NewAliases) {
                $FormattedAlias = $Alias 
                
                # Ensure alias is in the correct 'smtp:alias@domain.com' format
                if ($FormattedAlias -notlike "smtp:*") {
                    $FormattedAlias = "smtp:$FormattedAlias"
                }
                
                # *** NEW DUPLICATE CHECK ***
                if ($AddressesToKeepList.Contains($FormattedAlias)) {
                    Write-Warning "  -> SKIPPING duplicate alias from CSV: '$FormattedAlias' is already present in the list of addresses to keep."
                } else {
                    Write-Host "  -> Adding alias: $FormattedAlias" -ForegroundColor DarkYellow
                    $AddressesToKeepList.Add($FormattedAlias)
                }
            }
        }

        # Define the new primary and its duplicate alias
        # This was causing issues if both uppercase and lowercase (alias) versions existed, thanks MS365!
        $NewPrimarySMTP = "SMTP:$($Email)"
        $DuplicateAlias = "smtp:$($Email)" 

        # Remove the duplicate lowercase alias if it exists
        if ($AddressesToKeepList.Contains($DuplicateAlias)) {
            Write-Host "  -> Removing existing lowercase alias to avoid duplicate: $($DuplicateAlias)" -ForegroundColor DarkYellow
            $AddressesToKeepList.Remove($DuplicateAlias) | Out-Null
        }

        Write-Host "  -> Setting new Primary SMTP: $($Email)" -ForegroundColor DarkYellow
        # Add the correct new Primary SMTP (uppercase SMTP)
        $AddressesToKeepList.Add($NewPrimarySMTP)

        # Convert the List to a standard array *before* the Set-Mailbox command
        $FinalAddressesArray = $AddressesToKeepList.ToArray()

        Write-Host "  -> Final Address List for '$Identity' : $($FinalAddressesArray -join '; ')" -ForegroundColor DarkYellow

    }
    catch {
        Write-Warning "Failed to process identity '$Identity': $($_.Exception.Message)"
        $FormerAddressResults += [PSCustomObject]@{
            Identity = $Identity
            Aliases  = "ERROR: $($_.Exception.Message)"
        }
        continue # Skip to the next mailbox if address processing failed
    }
    ################################ [END OF PROXY ADDRESS PROCESSING BLOCK] ##########################

    ###################### [BOOLEAN CONVERSIONS AND PARAM SETUP] ######################
    # Convert the Account Status string to a boolean by using an explicit comparison. 
    # This is required for it to work later when passed via Update-MgUser

    $AccountStatus = ($AccountStatusString -eq "True")
    $AccountParams = @{
        accountEnabled = $AccountStatus
    }

    # Convert the value for HiddenFromAddressListsEnabled to a boolean
    $AccountHiddenFromGAL = ($AccountHiddenFromGALString -eq "True")

    # Password Profile (only used for the Destination tenant). We also set it so the user must change their password on next sign-in. We don't control this in the CSV as it is a pretty common expectation.
    $PasswordProfile = @{
        Password                      = $TemporaryPassword
        ForceChangePasswordNextSignIn = $true
    }
    ###################### [END OF BOOLEAN CONVERSIONS AND PARAM SETUP] ######################

    ############################### [MAIN USER UPDATE LOGIC BLOCK] ##############################
    # Inside the foreach loop, after setting $Identity and $Email:
    try {
        # Get the user object using the current identifier ($Identity, which is the current email/UPN)
        $UserObject = Get-MgUser -UserId $Identity -Property UserPrincipalName, Id
        #$CurrentUPN = $UserObject.UserPrincipalName, this is what I was using, but it was not reliable due to all the changes we're making at once
        $ObjectID = $UserObject.Id
    }
    catch {
        Write-Error "FAILURE: Could not retrieve user object for '$Identity'. Error: $($_.Exception.Message)"
        continue
    }
    
    if ([string]::IsNullOrEmpty($Identity) -or [string]::IsNullOrEmpty($Email)) {
        Write-Warning "Skipping row: Identity or Email is missing in a row. Please check the CSV file."
        continue
    }
    else {
        if ($Dryrun) {
            Write-Host "================= DISPLAYING USER CHANGES ================" -ForegroundColor Cyan
            Write-Host "[DRY RUN] No changes will be made for '$Identity'." -ForegroundColor Cyan
            Write-Host "[DRY RUN] Processing User: $Identity ..." -ForegroundColor Cyan
        }
        else {
            Write-Host "================= MIGRATION MODE ENABLED ================" -ForegroundColor Magenta
            Write-Host "Processing User: $Identity ..." -ForegroundColor Magenta
        }
        # Change Account Status based on $AccountStatus in spreadsheet
        try {
            if (-not $Dryrun) {
                Write-Host "Changing account enabled status for '$Identity' to '$AccountStatus'" -ForegroundColor Magenta
                Update-MgUser -UserId $ObjectID -BodyParameter $AccountParams -ErrorAction Stop
                Write-Host "SUCCESS: Account status (sign-in ability) for '$Identity' has been updated to '$AccountStatus'" -ForegroundColor Green
            }
            else {
                Write-Host "[DRY RUN] Skipping changing account enabled status for '$Identity' to '$AccountStatus'" -ForegroundColor Cyan
            }
        }
        catch {
            Write-Error "FAILURE: Could not modify account status for '$Identity'. Error: $($_.Exception.Message)"
            continue
        }
        # Configure HiddenFromAddressListsEnabled based on spreadsheet value
        try {
            if (-not $Dryrun) {
                Write-Host "Setting 'HiddenFromAddressListsEnabled' to '$AccountHiddenFromGAL' for '$Identity'" -ForegroundColor Magenta
                Set-Mailbox -Identity $Identity -HiddenFromAddressListsEnabled $AccountHiddenFromGAL -ErrorAction Stop
                Write-Host "SUCCESS: 'HiddenFromAddressListsEnabled' set to '$AccountHiddenFromGAL' for '$Identity'" -ForegroundColor Green
            }
            else {
                Write-Host "[DRY RUN] Skipping setting 'HiddenFromAddressListsEnabled' to '$AccountHiddenFromGAL' for '$Identity'" -ForegroundColor Cyan
            }
        }
        catch {
            Write-Error "FAILURE: Could not set 'HiddenFromAddressListsEnabled' for '$Identity'. Error: $($_.Exception.Message)"
            continue
        }
        # Change the user identity (UPN) based on the spreadsheet
        try {
            if (-not $Dryrun) {
                Write-Host "Changing User Principal Name (UPN) for '$Identity' to '$Email'" -ForegroundColor Magenta
                Update-MgUser -UserId $ObjectID -UserPrincipalName $Email -Mail $Email -ErrorAction Stop
                Write-Host "SUCCESS: Changed UPN for '$Identity' to '$Email'" -ForegroundColor Green
            }
            else {
                Write-Host "[DRY RUN] Skipping changing User Principal Name (UPN) for '$Identity' to '$Email'" -ForegroundColor Cyan
            }
        }
        catch { 
            Write-Error "FAILURE: Could not change Email for '$Identity' to '$Email'. Error: $($_.Exception.Message)"
            continue
        }
        # Change the actual email address, e.g $PrimarySMTP and $WindowsEmailAddress values in Exchange.
        try {
            if (-not $Dryrun) {
                Write-Host "Changing email address from $Identity to $Email" -ForegroundColor Magenta
                Write-Host "  -> New Email Addresses will be: $($FinalAddressesArray -join '; ')" -ForegroundColor DarkYellow
                # Use the fully converted array $FinalAddressesArray
                Set-Mailbox `
                -Identity $Identity `
                -EmailAddresses $FinalAddressesArray `
                -WindowsEmailAddress $Email `
                -MicrosoftOnlineServicesID $Email `
                -ErrorAction Stop
                Write-Host "SUCCESS: Changed email address from '$Identity' to '$Email'" -ForegroundColor Green
            }
            else {
                Write-Host "[DRY RUN] Skipping changing email address from $Identity to $Email" -ForegroundColor Cyan
                Write-Host "[DRY RUN]  -> New Email Addresses would be: $($FinalAddressesArray -join '; ')" -ForegroundColor Cyan
            }
        } 
        catch {
            Write-Error "FAILURE: Could not change email address from '$Identity' to '$Email'. Error: $($_.Exception.Message)"
            continue
        }
        # We want to change the passwords for Users in the Destination Tenant only, so filter for this condition
        # Source account passwords are intentionally left alone in case there is a reason an IT Admin needs to work with a user to access data post-migration
        # As the accounts are typically blocked from sign-in (disabled) after migration (via the AccountEnabledAtSource column in the -UsersCsv spreadsheet), this is not a security concern
        if (($Target -eq "Destination") -and ($MailboxType -eq "User")) {
            try {
                if (-not $Dryrun) {
                    Write-Host "Setting the password for '$Identity' as it is a destination tenant" -ForegroundColor Magenta
                    Update-MgUser -UserId $ObjectID -PasswordProfile $PasswordProfile -ErrorAction Stop
                    Write-Host "SUCCESS: Set password for '$Identity' to '$TemporaryPassword'" -ForegroundColor Green
                }               
                else {
                    Write-Host "[DRY RUN] Skipping setting the password for '$Identity' as it is a destination tenant" -ForegroundColor Cyan
                    Write-Host "[DRY RUN] Would set password for '$Identity' to '$TemporaryPassword'" -ForegroundColor Cyan
                }
                
            }
            catch { 
                Write-Error "FAILURE: Could not set password for '$Identity'. Error: $($_.Exception.Message)"
                continue
            }
        }
        elseif (($Target -eq "Destination") -and ($MailboxType -eq "Shared")) {
            Write-Host "Skipping password set for '$Identity' as it is a '$MailboxType' mailbox." -ForegroundColor Yellow
            Write-Host "If this is a mistake, correct the CSV and re-run the script" -ForegroundColor Yellow
        }
    }

}

# Disconnect from Exchange Online and MG Graph and end the script
# This is necessary to clean things up in case the script is run multiple times in the same session
# Also prevents a weird issue where the session can become invalid if left open for too long

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Disconnect-MgGraph | Out-Null

# Dump the information about former proxy addresses to a CSV for record-keeping
# Should be useful if you need to re-add any removed aliases later on. Technically we should be able to add them to the DestinationAliases column in the CSV 
# and re-run the script to re-add them if needed, though I haven't tested it yet.

$FormerAddressResults | Export-Csv -Path "$FormerAddresses" -NoTypeInformation

# Notify the user and wrap up the script
Write-Host "`nA record of the former proxy addresses have been exported to '$FormerAddresses'." -ForegroundColor Yellow
Write-Host "If needed, these can be copied to the 'DestinationAliases' column in the CSV and re-added. See the readme.md for more details." -ForegroundColor Yellow
Write-Host "`nScript execution complete." -ForegroundColor Green