# Microsoft 365 Migration Toolkit

## Overview

This repository serves as a collection of scripts to be used for Microsoft 365 migrations. At this time, the scripts are mostly focused on assisting with Microsoft to Microsoft migrations, which are common during Enterprise restructuring.

## Report Scripts

### User Report Script (M365_User_Report.ps1)
* The `M365_User_Report.ps1` script is useful to assist in data collection prior to the migration.
* It will collect inbox & OneDrive usage data, to help make business decisions (i.e: convert a dead, inactive account to a Shared Mailbox and move their data)
* *Global Reader* and *Reports Reader* is required for it to work properly

Usage: `./M365_User_Report.ps1 -TenantID YourTenantIDFromAzure -AdminUPN YourAdminAccount -DomainFilter contoso.com`

### Microsoft 365 Group Report Script (M365_Group_Report.ps1)
* Collects information on Microsoft 365 Groups and exports them to a CSV.
* Output includes the Group Name, Primary Email, whether the Group is a Microsoft Team, the size of the Mailbox in MB, the SharePoint Site URL (if defined), the amount of Owners, the amount of Members, the emails of the owners in a delimited format, the members of the group in a delimited format, and whether the group is Public or Private.

Usage: `./M365_Group_Report.ps1 -TenantID YourTenantIDFromAzure -AdminUPN YourAdminAccount -DomainFilter contoso.com`

### Microsoft 365 Resource Mailboxes Report (M365_Resource_Report.ps1)
* Collects information on Microsoft 365 shared mailboxes and room resource mailboxes, filtered by domain.
* Output includes the Primary SMTP Address (email) of the shared mailbox/resource group, the type of mailbox (Room or Shared), the size of the mailbox in MB, any email aliases, and the legacy exchange domain name.

Usage: `./M365_Resource_Report.ps1 -TenantID YourTenantIDFromAzure -AdminUPN YourAdminAccount -DomainFilter contoso.com`

## Migration Scripts

### Migration "Cutover" Script (M365_Migrate_Tenant.ps1)
* The `M365_Migrate_Tenant.ps1.ps1` script is meant to help with assisting in Microsoft 365 to Microsoft 365 migrations, ideally while using MigrationWiz, Movebot.io, or a similar tool. 
* It was created after assisting a large enterprise with multiple domain migrations and divison of assets. 
* The script will make changes against a `Source` and `Destination` tenant, based on values you provide in the `TenantsCsv` file.
* It is intended to be somewhat idempotent, so you can run it multiple times if needed, and the parameters defined in the CSV are the "source of truth" regarding state.

Usage: `./M365_Migrate_Tenant.ps1 -TenantsCsv TenantConfig.csv -UsersCsv UserInformation.csv -Target Source|Destination -TargetDomain contoso.com`

#### Parameters and Switches

The following parameters are required:
* `TenantsCSV` : The path to the CSV file you're using for your Tenant(s) configuration data.
* `UsersCSV` : The path to the CSV file you're using for your User data.
* `TargetDomain` : The domain you're migrating. This is used in some of the logic within the script for handling proxy addresses. Any addresses found with the domain are discarded, other addresses with other domains are untouched.
* `Target` : Used to pass either `Source` or `Destination` - tells the script which tenant you're intending to make changes to. Values from the provided `-UsersCsv` are pulled accordingly, see above.

The following switches are optional:
#### Switches
* `-Revert` : Used to "revert" changes, but will not solve for aliases. For the Source Tenant, it will swap `SourceEmail` with `PostMigrationSourceEmail`. It will look for the account named `PostMigrationSourceEmail` and (presumably) change it *back* to `SourceEmail`. For the Destination Tenant, it will swap `DestinationStagingEmail` with `DestinationPostMigrationEmail`. It will look for the account named `DestinationStagingEmail` and (presumably) change it *back* to `DestinationPostMigrationEmail`.
* `-Dryrun` : Extremely useful! Used to do a dry run of the script, and will output the proposed changes in console. **Highly recommend you run this once or twice prior to doing your migration.** It will also collect the Users' existing proxy addresses and export them to a CSV for you.

You will then see a .csv within the project directory named `FormerAddresses_domain.com_DATE.csv` once it completes.


# How to use these tools during your migration

1. Collect information on the domain you're migrating, running the various report scripts.
2. Discuss with stakeholders what needs to be migrated, archived, etc., based on the results of the reports.
2. Collect the information into a "source" CSV for the changes you want to make.
3. Use that CSV to perform the migration, which can be adapted for the Microsoft Migration Tool, MigrationWiz, and others.
4. Use the `M365_Migrate_Tenant.ps1` script to perform the "cutover" operations, in the `Source` and `Destination` tenants

In the example below, let's assume you are migrating from a Microsoft 365 Tenant to another Microsoft 365 Tenant. As apart of the migration, you're also standardizing to a new format, e.g switching from `firstname@domain.com` to `firstname.lastname@domain.com`.

## Preparing for a Migration

You will need to create (2) CSV files for using the script, with the following values in each. They can be assembled using the various report scripts. At some point in the future, we may combine all of the scripts to make one file, but for now it is a manual process. There are likely business decisions to be made during the migration, so it likely makes sense in many cases to perform a review of the reports and assemble the CSVs.

In the CSV you are creating for the cutover script, you will need the following columns:

* `Users.csv` : A CSV file with all of the user data. Must contain the following columns:
  * `SourceHideFromGAL` : Toggle on whether the user should be hidden from the GAL in the *source* tenant. Set to `TRUE` to hide, set to `FALSE` to make the account viewable.
  * `SourceEmail` : The user identity we're modifying or making changes against. This is typically the user's `UserPrinicipalName` or `UPN`
  * `PostMigrationSourceEmail` : The intended email for the identity, post-migration. The primary email for the user will be changed to this when running the script.
  * `DestinationHideFromGAL` : Toggle on whether the user should be hidden from the GAL in the *destination* tenant. Set to `TRUE` to hide, set to `FALSE` to make the account viewable.
  * `DestinationAliases` : The intended alias for the identity on the destination tenant. Useful when switching from `firstname@domain.com` to `firstname.lastname@domain.com`, and you need to keep emails sent to `firstname@domain.com`. Should support multiple aliases separated by a `;`, e.g: `firstname@domain.com; lastname@domain.com; bob@domain.com`, but this hasn't been tested yet.
  * `DestinationStagingEmail` : The "staging" email in the destination tenant for the user. This is often a temporary identity, such as an `domain.onmicrosoft.com` account. It is typically the identity you migrate data to in MigrationWiz.
  * `PostMigrationDestinationEmail` : The intended primary email for the user in the destination tenant, post-migration. Using the current example, this would be `firstname.lastname@domain.com`
  * `DestinationPassword` : The password that should be assigned to the account so they can login post-migration.
  * `AccountEnabledAtSource` : Controls whether or not the identity at the Source is enabled. Set to `$true` or `$false`. Is evaluated each time the script is ran.
  * `AccountEnabledAtDestination` : Controls whether or not the identity at the Destination is enabled. Set to `$true` or `$false`. Is evaluated each time the script is ran.

* `Tenants.csv` : A CSV file with the Tenant information, used for connecting the PowerShell modules for `MGGraph` and `ExchangeOnlineManagement`.
  * `SourceTenantID` : The `TenantID` from the Source tenant, found in the Azure portal.
  * `SourceTenantAdmin` : The `UserPrincipalName` of your Administrator account in the Source tenant, e.g `admin@domain.com`.
  * `DestinationTenantID` : The `TenantID` from the Destination tenant, found in the Azure portal.
  * `DestinationAdmin` : The `UserPrincipalName` of your Administrator account in the Destination tenant, e.g `admin@otherdomain.com`.

**It is highly recommended you fill out the User data with 1-2 test accounts and run the script with those first, so you can get a feel for how the script operates.** I cannot emphasize this enough. The script does
have some "safety" built in, but it is up to you to use this responsibly.**It is important to understand that the script will iterate through each row in the User CSV and make changes, and that you should test and become familiar with the script before running it across multiple accounts, or an entire organization... you've been warned.**


# Planning a Migration - Scenario Example

Let's use a common scenario as an example of how to use this toolkit:
* You're tasked with completing a migration in 1 week.
* You need to pull data on the accounts to make informed decisions (e.g convert inactive User Accounts to Shared Mailboxes)
* You're using MigrationWiz and need to pre-stage the data, then run the final delta sync overnight when the domain is cutover.

## Step 1: Collect your Data

Run the report scripts and collect the data about your resources. See the Report Scripts section for more information on usage and expected output. Assemble all your findings into a CSV using the provided templates.

## Step 2: Complete the Cutover & Move the Domain

Assuming you are doing a cutover migration in the evening, here is how this would work:
1. Run `M365_Migrate_Tenant.ps1 -UsersCSV Users.csv -TenantsCSV Tenants.CSV -Target Source -TargetDomain domain.com` 
2. Remove domains from the old tenant, and verify domain in the new tenant
3. At this point, users will be on `domain.onmicrosoft.com` in source tenant
4. In MigrationWiz, switch Source using the Change Domain tool to change to `domain.onmicrosoft.com`
5. At this point, users in the Target Tenant will still be set to using their temporary domain (`firstname.lastname@domain.com`). We need to add their old email address as an alias (`firstname@domain.com`) and convert them to use the new domain (`firstname.lastname@domain.com`), keeping the domain prefix (mail nickname)
6. Using the `ConvertEmails.ps1` script, convert the users in the new tenant from their staged account to their final account.
7. Complete Final sync

If you need to activate/deactivate an account or hide/unhide an account from the GAL, simply change those values and re-run the script as necessary.