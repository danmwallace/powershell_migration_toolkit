<# 
.SYNOPSIS
    This script sets the 'HiddenFromAddressListsEnabled' property for mailboxes in either the Source or Destination tenant based on a provided CSV file. 
    This is a lightweight script in case for some reason you need to hide mailboxes from the Global Address List (GAL) without making additional changes.
.DESCRIPTION
    The script imports mailbox data from a CSV file and updates the 'HiddenFromAddressListsEnabled
    property for each mailbox in the specified tenant (Source or Destination) according to the values in the CSV.

    Ideally we would reference the same Users.csv file used for the main migration script to ensure consistency, but you can create a separate CSV if needed.
.PARAMETER UsersCsv
    The path to the CSV file containing mailbox data with columns for SourceEmail, DestinationEmail,
    SourceHideFromGAL, and DestinationHideFromGAL, and likely additional columns as well (see readme.md for details).
.PARAMETER TenantsCsv
    The path to the CSV file containing tenant admin details with columns for SourceTenantID,
    DestinationTenantID, SourceTenantAdmin, and DestinationTenantAdmin.
.PARAMETER Target
    Specifies whether to update the 'Source' or 'Destination' tenant.
.EXAMPLE
    .\HideFromGAL.ps1 -UsersCsv "C:\Path\To\Users.csv" -TenantsCsv "C:\Path\To\Tenants.csv" -Target "Source"
    This command sets the 'HiddenFromAddressListsEnabled' property for mailboxes in the Source tenant based on the provided Users.csv file.
.NOTES
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$UsersCsv,

    [Parameter(Mandatory=$true)]
    [string]$TenantsCsv,

    [Parameter(Mandatory=$true)]
    [string]$Target
)

############################# [IMPORT USER AND TENANT DATA] #############################
$MailboxData = Import-Csv -Path $UsersCsv
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
############################# [CONNECT TO TENANT] #############################
# Connect with MG Graph and ExchangeOnline(Management) before running the script
Connect-MGGraph -Scopes "User.ReadWrite.All" -TenantId $TenantID -NoWelcome
Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName

############################# [SET HIDDEN FROM GAL BASED ON CSV] #############################
foreach ($Mailbox in $MailboxData) {
    if ($Target -eq "Source") {
        $Identity = $Mailbox.SourceEmail
        $HiddenFromGALValue = $Mailbox.SourceHideFromGAL
    }
    elseif ($Target -eq "Destination") {
        $Identity = $Mailbox.DestinationEmail
        $HiddenFromGALValue = $Mailbox.DestinationHideFromGAL
    }
    # Convert the value for HiddenFromAddressListsEnabled to a boolean
    $AccountHiddenFromGAL = ($HiddenFromGALValue -eq "True")
    Write-Host "Setting 'HiddenFromAddressListsEnabled' to '$AccountHiddenFromGAL' for '$Identity'" -ForegroundColor White
    Set-Mailbox -Identity $Identity -HiddenFromAddressListsEnabled $AccountHiddenFromGAL -ErrorAction Stop
}