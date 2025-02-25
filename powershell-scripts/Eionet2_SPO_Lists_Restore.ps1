# -------------------------------------------------------------------------
# This PS script is used to restore the SPO lists used by the Eionet 2 apps.
# It restores the content of each SPO list from the backup .xml file.
# The script uses the PnP PowerShell module - see https://pnp.github.io/powershell/
# Last update: April 2023
# -------------------------------------------------------------------------

# IMPORTANT! YOU MUST UPDATE THIS WITH YOUR SPO SITE URL (where the Eionet 2 lists are stored)
$SPO_SITE_PATH = "https://2xkk2b.sharepoint.com/sites/Testsite" # (Daniel's tenant)
#$SPO_SITE_PATH = "https://7lcpdm.sharepoint.com/sites/EIONETPortal" # (Mihai's tenant)
#$SPO_SITE_PATH = "https://<...>.sharepoint.com" # (EEA's tenant)

$SPO_LISTS = @()
# IMPORTANT! If you need to add/remove a list, make sure you use its "Display name" property (not its "Name" property).
# IMPORTANT! The order in whcih the lists are restored is very important!
#            You have to make sure you handle properly the dependencies between lists when you restore them.
$SPO_LISTS += "ConfigurationList"
$SPO_LISTS += "MappingList"
$SPO_LISTS += "OrganisationList"
$SPO_LISTS += "UsersList"
$SPO_LISTS += "EventsList"
$SPO_LISTS += "Events Participants"
$SPO_LISTS += "ConsultationsList"
#$SPO_LISTS += "LoggingList" #This is a huge list. During testing you might want to skip it by commenting this line.
<#
# Some local small test lists - to be removed in the script release version
$SPO_LISTS += "EionetUsersList"
$SPO_LISTS += "Eionet-Organizations-List"
#>

# Ensure that the PnP PS module is installed in your system.
# See https://pnp.github.io/powershell/
#Write-Output "Uninstall the PnP module..."
#Uninstall-Module -Name PnP.PowerShell -Force
Write-Output "Ensuring that the PnP module is installed..."
If (-not(Get-Module -Name PnP.PowerShell -ListAvailable)) {
    Write-Output "The PnP module was not found - installing..."

    # We need to install the PnP module
    # Check first if script is executed under elevated permissions - Run as Administrator
    If (-not([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
        Write-Output "Please run this script in elevated mode (Run as Administrator)!"
        Write-Output "Exiting..."
        Exit
    }

    # There are potential breaking changes in the latest 2+ versions, released after March 2023 (lost compatibility with PS 5.1 etc).
    # Therefore, we will use for now the previous stable major release - 1.12.
    #Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -Confirm:$False -RequiredVersion 1.12.0
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -RequiredVersion 1.12.0

    # Ensure that the PnP module functions have access to the tenant SPO resources.
    # See https://pnp.github.io/powershell/cmdlets/Register-PnPManagementShellAccess.html
    # See also https://dev.to/svarukala/introducing-the-new-pnp-powershell-based-on-net-core-3-1-and-learn-how-it-s-authentication-works-pn7
    # This PS command will create an Azure AD Enterprise Application (a service principal) with an ID (31359c7f-bd7e-475c-86db-fdb8c937548e)
    # and the name of this application is "PnP Management Shell".
    # You can navigate to your Azure Portal > Azure Active Directory > Enterprise Applications. You can see the app in there.
    # If you are not an administrator that can consent Azure AD Applications, use the -ShowConsentUrl option.
    # It will ask you to log in and provides you with an URL you can share with a person with appropriate access rights
    # to provide consent for the organization.
    # NOTE: You don’t need to be a Tenant Admin to use the PnP.PowerShell cmdlets.
    # You don’t even need to be a SharePoint Admin or a site collection admin.
    # There are plenty of cmdlets you can run. However, before you can run the most import PnP cmdlet of all, Connect-PnPOnline,
    # the PnP Azure Application has to be registered in your tenant by a tenant admin.
    Write-Output "Ensure that the PnP module functions have access to the required tenant resources..."
    Register-PnPManagementShellAccess
} Else {
    Write-Output "The PnP module is already installed."
    #Write-Output "Checking for PnP updates..."
    #Update-Module -Name PnP.PowerShell -Scope CurrentUser -Confirm
}

# Connect to the SPO site.
# User will be asked for credentials if not already authenticated.
# See https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html
Write-Output (-join("Connecting to the SPO site: ", $SPO_SITE_PATH, "..."))
# "-Interactive": Connects to the Azure AD using interactive login, allowing you to authenticate using MFA.
Connect-PnPOnline -Url $SPO_SITE_PATH -Interactive
# If you have saved your login credentials in the browser, you can let PowerShell fetch them with the command "-UseWebLogin"
#Connect-PnPOnline -Url $SPO_SITE_PATH -UseWebLogin
#Connect-PnPOnline -Url $SiteURL -Credentials (Get-Credential)

If (-not(Get-PnPContext)) {
    Write-Output "Error connecting to SharePoint Online, unable to establish context!"
    Write-Output "Exiting..."
    Exit
}

$restored_lists = @()
# Process each predefined list
Foreach ($list_name in $SPO_LISTS) {
    Write-Output (-join("Restoring list ", $list_name, "..."))

    # Check if the list backup file exists
    $backup_list_name = -join($list_name, ".xml")
    If (-not(Test-Path $backup_list_name)) {
        Write-Output (-join("`tThe backup file ", $backup_list_name, " does not exist! Skipping this list restore..."))
        Continue
    }

    # Check if list already exists; it yes, it needs to be removed first.
    $list = Get-PnPList -Identity $list_name -ErrorAction SilentlyContinue
    If ($List -ne $Null) {
        # Delete the existing list. The user will be asked to confirm the list removal.
        Write-Output (-join("`tA list with name ", $list_name, " already exists! Please remove it."))
        # Remove the -Recycle in order to not send the list to the Recycle Bin
        Remove-PnPList -Identity $list_name -Recycle
    }

    # Check if the list was deleted; if not, skip it from being restored.
    $list = Get-PnPList -Identity $list_name -ErrorAction SilentlyContinue
    If ($List -ne $Null) {
        Write-Output (-join("`tA list with name ", $list_name, " still exists! Skipping this list restore..."))
        Continue
    }

    # Restore list from the backup file
    Try {
        Invoke-PnPSiteTemplate -Path $backup_list_name -Handlers Lists
        $restored_lists += $list_name
    } Catch {
        Write-Output (-join("`tError restoring the list ", $list_name, " from the file: ", $backup_list_name))
        Write-Output (-join("`t", $_.Exception.Message))
        Write-Output "`tSkipping this list..."
        Continue
    }
}

# Output the results
Write-Output "--------------------------------------"
Write-Output (-join($restored_lists.Count, " out of ", $SPO_LISTS.Count, " lists were restored."))
If ($restored_lists.Count -ne 0) {
    Write-Output "The following lists were restored:"
    Foreach ($list_name in $restored_lists) {
        $list_rows_count = (Get-PnPListItem -List $list_name).Count
        Write-Output (-join("`t", $list_name, " (", $list_rows_count, " rows)"))
    }
}
$not_restored_lists = $SPO_LISTS | Where {$restored_lists -NotContains $_}
If ($not_restored_lists.Count -ne 0) {
    Write-Output "The following lists were NOT restored:"
    Foreach ($list_name in $not_restored_lists) {
        Write-Output (-join("`t", $list_name))
    }
}

# Disconnects the current connection and clears its token cache.
# It will require you to build up a new connection again using Connect-PnPOnline in order to use any of the PnP PowerShell cmdlets.
# You will have to reauthenticate.
#Disconnect-PnPOnline
