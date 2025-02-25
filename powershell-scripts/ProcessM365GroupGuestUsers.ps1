#Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
<#
  Scope:
    - For every user member of a M365 group, get its properties and the sign-in activity. Save this data in a CSV file.
    - For every guest member of a M365 group, get its properties and do a DisplayName rename if the following conditions are met:
      * The Department is not "Eionet"
      * The Country (whic is obtained from the Mail address) is not null.
      * The DisplayName does not end with ' (XX)', where XX is the Country ISO-2 code.
      The rule for rename is: 'DisplayName' becomes 'DisplayName (XX)', where XX is the Country ISO-2 code.
      Note; The update rule should take into consideration the already updated users and checking if the country code is valid.
  Resources:
    https://learn.microsoft.com/en-us/microsoft-365/enterprise/connect-to-microsoft-365-powershell
    https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands
    https://learn.microsoft.com/en-us/graph/throttling-limits?view=graph-rest-1.0#identity-and-access-service-limits
#>

# Stop On Errors
$ErrorActionPreference = 'Stop'

#################################################################################
# DEFINE HELPER FUNCTIONS
#################################################################################

function ForegroundColorGreen
{
    process { Write-Host $_ -ForegroundColor Green }
}

function ForegroundColorRed
{
    process { Write-Host $_ -ForegroundColor Red }
}

function ForegroundColorYellow
{
    process { Write-Host $_ -ForegroundColor Yellow }
}


# Clear PowerShell console window
Clear-Host

#################################################################################
# CHECK Microsoft.Graph MODULE INSTALLATION
#################################################################################

#if (!(Get-InstalledModule -Name Microsoft.Graph))
if (!(Get-Module -ListAvailable -Name Microsoft.Graph))
{
    Write-Output "Microsoft.Graph module is not loaded. Installing..."
    Install-Module -Name Microsoft.Graph -Force -Verbose
    #Install-Module -Name Microsoft.Graph -Force -Scope CurrentUser -Verbose
    #Install-Module -Name Microsoft.Graph.Beta -Force -Scope CurrentUser -Verbose
    Write-Output "Please re-run the script. Exiting."
    exit
} else {
    Write-Output "Microsoft.Graph module is loaded."
}

# Uncomment this line if you want to update the Microsoft.Graph module if there is a new version available
#Update-Module -Name Microsoft.Graph -Force -Verbose

#################################################################################
# DEFINE CONSTANTS AND PARAMETERS
#################################################################################

# https://entra.microsoft.com/#home
$TenantId = "" # EEA's tenant
$TenantId = "" # Local dev tenant

#$M365GroupId = "6327153c-b9c0-4158-ae98-6c84267a0176"  # EPANET
$M365GroupId = "4aa828ba-f097-4d7c-ac60-5407e9856060"  # Local dev tenant group "All company"
<# 
# Alternative way to get the M365 group Id:
$M365GroupName = "EPANET"
$M365GroupName = "All Company"
$groupId = (Get-MgGroup -Filter "DisplayName eq '$M365GroupName'").Id
#>


<# 
Define the permissions needed.
Get-MgGroupMember:
    GroupMember.Read.All
    GroupMember.ReadWrite.All
    Group.ReadWrite.All
    Group.Read.All
    Directory.Read.All
Note: You can use this PS command: >(Find-MgGraphCommand -Command Get-MgGroupMember).permissions to find the needed permissions

Update-MgUser:
    User.ReadWrite.All
    User.ReadWrite
    User-PasswordProfile.ReadWrite.All
    User-Mail.ReadWrite.All
    Directory.ReadWrite.All
    DeviceManagementServiceConfig.ReadWrite.All
    DeviceManagementManagedDevices.ReadWrite.All
    DeviceManagementConfiguration.ReadWrite.All
    User.ManageIdentities.All
    User.EnableDisableAccount.All
    User-Phone.ReadWrite.All
    DeviceManagementApps.ReadWrite.All  
Note: You can use this PS command: >(Find-MgGraphCommand -Command Update-MgUser).permissions to find the needed permissions
#>
$RequiredScopes = @("User.Read.All", "AuditLog.Read.All", "GroupMember.Read.All", "GroupMember.ReadWrite.All", "Group.ReadWrite.All", "Group.Read.All")
$RequiredScopes += @("Directory.Read.All")
$RequiredScopes += @("User.ReadWrite.All", "User.ReadWrite", "User-PasswordProfile.ReadWrite.All", "User-Mail.ReadWrite.All", "Directory.ReadWrite.All")
$RequiredScopes += @("DeviceManagementServiceConfig.ReadWrite.All", "DeviceManagementManagedDevices.ReadWrite.All", "DeviceManagementConfiguration.ReadWrite.All")
$RequiredScopes += @("User.ManageIdentities.All", "User-Phone.ReadWrite.All", "DeviceManagementApps.ReadWrite.All")

# Save the current date stamp
$LogDate = Get-Date -f yyyy-MM-dd-hh-mm

# Define the CSV export files
$OutputGroupUsers_CSVPath = ".\GroupUsers_$LogDate.csv"
$OutputGroupGuestUsersDisplayNameRename_CSVPath = ".\GroupGuestUsers_DisplayNameRename_$LogDate.csv"

#################################################################################
# PREPARE LIST OF VALID COUNTRIES ISO-2 CODES
#################################################################################

$AllCultures = [System.Globalization.CultureInfo]::GetCultures([System.Globalization.CultureTypes]::SpecificCultures)
$ValidCountries = @();
$AllCultures | % {
    $RegionInfo = New-Object System.Globalization.RegionInfo $PsItem.name;
    $ValidCountries += $RegionInfo.TwoLetterISORegionName
}
$ValidCountries = $ValidCountries | Select -Unique | Sort-Object
$ValidCountries += "EU"


#################################################################################
# CONNECT TO TENANT
#################################################################################

Write-Output "M365 Group Id to process: $M365GroupId"
Write-Output "Connect to MS Graph..."
<# 
Connect using Interactive mode (delegated access).
If you’ve never signed in with the Graph SDK before, the SDK creates an enterprise app called Microsoft Graph Command Line Tools
with an AppId of 14d82eec-204b-4c2f-b7e8-296a70dab67e and requests a limited set of permissions.
If you’re an administrator, you can grant consent for these permissions on behalf of the organization.
#>
Connect-MgGraph -TenantId $TenantId -Scopes $RequiredScopes -NoWelcome 

<# 
Note: The only resolution for an over-permissioned service principal is its removal and recreation, 
at which time an administrator can grant consent for limited permissions to the new service principal. 
Here’s how to remove the service principal using Graph SDK cmdlets (naturally):
$Sp = Get-MgServicePrincipal -Filter "AppId eq '14d82eec-204b-4c2f-b7e8-296a70dab67e'"
Remove-MgServicePrincipal -ServicePrincipalId $Sp.Id
#>

# Display connection information
$ConnectionDetails = Get-MgContext
$Scopes = $ConnectionDetails | Select -ExpandProperty Scopes
$Scopes = $Scopes -Join ", "
$OrgName = (Get-MgOrganization).DisplayName
$CurrentContext = Get-MgContext
$CurrentUser = $CurrentContext.Account
Write-Output "+-------------------------------------------------------------------------+"
Write-Output "Microsoft Graph Connection Information"
Write-Output "+-------------------------------------------------------------------------+"
Write-Output "Connected to tenant $TenantId, Org: $OrgName"
Write-Output "Current user: $CurrentUser"
Write-Output "+-------------------------------------------------------------------------+"
Write-Output ("The following permission scopes are used: {0}" -f $Scopes)
Write-Output "+-------------------------------------------------------------------------+"
Write-Output ""

#################################################################################
# RETRIVE USER MEMBERS OF THE GROUP
#################################################################################

$M365GroupName = (Get-MgGroup -Filter "Id eq '$M365GroupId'").DisplayName
Write-Output "Get all members of the M365 group: $M365GroupName ($M365GroupId)"
# Note: A group can have users, organizational contacts, devices, service principals and other groups as members.
#$groupMembers = Get-MgGroupMember -GroupId $M365GroupId -All -PageSize 999 | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' }
#$groupMembers = Get-MgGroupMember -GroupId $M365GroupId -All -PageSize 999
# https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.groups/get-mggroupmemberasuser?view=graph-powershell-1.0
$groupMembers = Get-MgGroupMemberAsUser -GroupId $M365GroupId -All -PageSize 999
$groupMembersCount = $groupMembers.Count
Write-Output "Found $groupMembersCount direct users in the group."

if ($groupMembersCount -eq 0) {
    Write-Host "Didn't found any members in the group. Exiting."
    exit
}
Write-Output ""


#################################################################################
# PROCESS USERS
#################################################################################

Write-Output "Find users information..."
$users = @()
$usersToUpdate = @()
$usersNotUpdated = @()
foreach ($member in $groupMembers) {
    #$properties = $member.AdditionalProperties
    #if ($properties['@odata.type'] -eq '#microsoft.graph.user') {
        #-Filter "signInActivity/lastSignInDateTime lt $([datetime]::UtcNow.AddDays(-100).ToString("s"))Z"
        $groupUser = Get-MgUser -UserID $member.Id -Property DisplayName, UserPrincipalName, Mail, Country, UserType, Department, SignInActivity
        
        $DisplayName       = $groupUser.DisplayName
        $UserPrincipalName = $groupUser.UserPrincipalName
        $Mail              = $groupUser.Mail
        $Country           = $groupUser.Country
        $UserType          = $groupUser.UserType
        $Department        = $groupUser.Department
        $LastSignIn        = $groupUser.SignInActivity.LastSignInDateTime
        $LastNonInteractiveSignIn = $groupUser.SignInActivity.LastNonInteractiveSignInDateTime
        $Id                = $groupUser.Id

        Write-Output "  DisplayName      : $DisplayName"
        Write-Output "  UserPrincipalName: $UserPrincipalName"
        Write-Output "  Mail             : $Mail"
        Write-Output "  Country          : $Country"
        Write-Output "  UserType         : $UserType"
        Write-Output "  Department       : $Department"
        Write-Output "  LastSignIn       : $LastSignIn"
        Write-Output "  LastNonInteractiveSignIn: $LastNonInteractiveSignIn"
        Write-Output "  Id       : $Id"

        $objUser = [PSCustomObject] @{
            DisplayName       = $DisplayName
            UserPrincipalName = $UserPrincipalName
            Mail              = $Mail
            Country           = $Country
            UserType          = $UserType
            Department        = $Department
            LastSignIn        = $LastSignIn
            LastNonInteractiveSignIn = $LastNonInteractiveSignIn
            Id                = $id
        }
        $users += $objUser;

        $CountryFromMail = ""
        if ($Mail -ne $null) {
            $CountryFromMail = $Mail.Split(".")[-1].ToUpper()
        }
        $IsDisplayNameAlreadyUpdated = $false
        if ($CountryFromMail.length -eq 2) {
            if ($DisplayName.EndsWith(" (" + $CountryFromMail + ")")) {
                $IsDisplayNameAlreadyUpdated = $true
                Write-Output "  DisplayName is already updated for user $Mail" | ForegroundColorYellow
            }
        }
        $IsCountryCodeValid = $false
        if ($CountryFromMail.length -eq 2) {
            if ($ValidCountries.Contains($CountryFromMail)) {
                $IsCountryCodeValid = $true
            } else {
                 Write-Output "  This is not a valid ISO-2 country code: $CountryFromMail" | ForegroundColorYellow
            }
        }

        # Conditions to update the DisplayName of the user:
        #  - Department is not 'Eionet'
        #  - User is a 'Guest'
        #  - There is valid country code in the email address
        #  - The DisplayName is not already updated
        #  - The country code is valid

        if (($Department -ne 'Eionet') -and ($UserType -eq 'Guest') -and ($CountryFromMail.length -eq 2) -and
            ($IsDisplayNameAlreadyUpdated -ne $true) -and ($IsCountryCodeValid -eq $true)) {
            # The DisplayName will can be changed
            $objUserToUpdate = [PSCustomObject] @{
                DisplayName    = $DisplayName
                Mail           = $Mail
                NewDisplayName = $DisplayName + " (" + $CountryFromMail + ")"
                Id             = $Id
            }
            $usersToUpdate += $objUserToUpdate;
        } else {
            # The DisplayName will not be changed
            $Comment = ""
            if ($Department -eq 'Eionet') {
                $Comment = "User department is Eionet."   
            }
            if ($UserType -ne 'Guest') {
                $Comment = "User is not a Guest."   
            }
            if ($CountryFromMail.length -ne 2) {
                $Comment = "Cannot get an ISO-2 country code from the email address."   
            }
            if ($IsDisplayNameAlreadyUpdated -eq $true) {
                $Comment = "The DisplayName is already updated."
            }
            if ($IsCountryCodeValid -eq $false) {
                $Comment = "Cannot get a valid ISO-2 country code from the email address."
            }
            $objUserNotUpdated = [PSCustomObject] @{
                DisplayName  = $DisplayName
                Mail         = $Mail
                Comment      = $Comment
            }
            $usersNotUpdated += $objUserNotUpdated;
        }
    #} else {
    #    $groupUser = Get-MgUser -UserID $member.Id | Select DisplayName
    #    $groupUserDisplayName = $member.DisplayName
    #    Write-Output "  User $groupUserDisplayName is not of type '#microsoft.graph.user'. Skipping." | ForegroundColorYellow
    #}
    Write-Output ""
}

if ($users.Count -eq 0) {
    Write-Host "Didn't found any users to process. Exiting."
    exit
}
Write-Output ""

#################################################################################
# SAVE ALL USERS IN A CSV FILE
#################################################################################

Write-Output "Save users information in a CSV file..."
if (Test-Path $OutputGroupUsers_CSVPath) {
    Write-Output "Delete file $OutputGroupUsers_CSVPath..."
    Remove-Item $OutputGroupUsers_CSVPath -Verbose
}
$users | Select-Object DisplayName, UserPrincipalName, Mail, Country, UserType, Department, LastSignIn, LastNonInteractiveSignIn |
         Sort-Object -Property LastSignIn -Descending |
         Export-Csv -Path $OutputGroupUsers_CSVPath -NoTypeInformation -Encoding UTF8
Write-Output "Users information were exported to $OutputGroupUsers_CSVPath."
Write-Output ""

if ($usersToUpdate.Count -eq 0) {
    Write-Host "Didn't found any users to update. Exiting."
    exit
}
Write-Output ""

#################################################################################
# SHOW THE USERS WHO WILL BE UPDATED AND SAVE TO CSV
#################################################################################

Write-Output "Suggested changes for users:"
$usersToUpdate | Select-Object DisplayName, Mail, NewDisplayName | Format-Table -AutoSize
Write-Output "Save suggested changes in a CSV file..."
if (Test-Path $OutputGroupGuestUsersDisplayNameRename_CSVPath) {
    Write-Output "Delete file $OutputGroupGuestUsersDisplayNameRename_CSVPath..."
    Remove-Item $OutputGroupGuestUsersDisplayNameRename_CSVPath -Verbose
}
$usersToUpdate | Select-Object DisplayName, Mail, NewDisplayName |
                 Export-Csv -Path $OutputGroupGuestUsersDisplayNameRename_CSVPath -NoTypeInformation -Encoding UTF8
Write-Output "Suggested users changes were exported to $OutputGroupGuestUsersDisplayNameRename_CSVPath."
Write-Output ""


#################################################################################
# SHOW THE USERS WHO WILL NOT BE UPDATED
#################################################################################

Read-Host -Prompt "Press any key to display the users who will not be updated..." | Out-Null


Write-Output "The following users will not be updated:" | ForegroundColorYellow
#$usersNotUpdated | Select-Object DisplayName, Mail, Comment | Format-Table -AutoSize | ForegroundColorYellow
$usersNotUpdated | Select-Object DisplayName, Mail, Comment | Format-Table -AutoSize

$UpdateCount = $usersToUpdate.Count
if ($UpdateCount -eq 0) {
    Write-Host "Didn't found any users to update. Exiting."
    exit
}
Write-Output ""

#################################################################################
# UPDATE THE USERS
#################################################################################

$ConfirmMsg = "You are about to change the DisplayName for the selected users. If you want to continue enter 'yes':"
Write-Host -NoNewline $ConfirmMsg | ForegroundColorRed
$Response = Read-Host
if ($Response -ne "yes") {
    Write-Host "Exiting."
    exit
}
Write-Output ""

# Update users
Write-Output "$UpdateCount users will be updated."
foreach ($user in $usersToUpdate) {
    $Id = $user.Id
    $Mail = $user.Mail
    $NewDisplayName = $user.NewDisplayName
    Write-Output "Update user $Mail to set DisplayName to $NewDisplayName..."

    # REMOVE THE PARAMETER  -WhatIf TO ACTUALLY UPDATE THE USER
    #Update-MgUser -UserId $Id -DisplayName $NewDisplayName -WhatIf
    Update-MgUser -UserId $Id -DisplayName $NewDisplayName

    Write-Output "New DisplayName for user $Mail is '$NewDisplayName'." | ForegroundColorGreen
}
Write-Output "$UpdateCount users were updated."
Write-Output ""

Disconnect-MgGraph
