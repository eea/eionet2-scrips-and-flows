# -------------------------------------------------------------------------
# This PS script is used to restore the SPO lists used by the Eionet 2 apps.
# It restores the content of each SPO list from the backup .xml file.
# The script uses the PnP PowerShell module - see https://pnp.github.io/powershell/
# Last update: April 2023
# -------------------------------------------------------------------------

# IMPORTANT! YOU MUST UPDATE THIS WITH YOUR SPO SITE URL (where the Eionet 2 lists are stored)
$SPO_SITE_PATH = "https://2xkk2b.sharepoint.com/sites/Testsite" # (Daniel's tenant)
#$SPO_SITE_PATH = "https://7lcpdm.sharepoint.com/sites/EIONETPortal" # (Mihai's tenant)
#$SPO_SITE_PATH = "https://EEA1.sharepoint.com" # (EEA's tenant)

$SPO_LISTS = @()
# IMPORTANT! If you need to add/remove a list, make sure you use its "Display name" property (not its "Name" property).
# IMPORTANT! The order in whcih the lists are restored is very important!
#            You have to make sure you handle properly the dependencies between lists when you restore them.
#$SPO_LISTS += "ConfigurationList"
#$SPO_LISTS += "MappingList"
#$SPO_LISTS += "OrganisationList"
#$SPO_LISTS += "UsersList"
#$SPO_LISTS += "EventsList"
#$SPO_LISTS += "Events Participants"
#$SPO_LISTS += "ConsultationsList"
$SPO_LISTS += "LoggingList" #This is a huge list. During testing you might want to skip it by commenting this line.
<#
# Some local test lists - to be removed in the script release version
$SPO_LISTS += "EionetUsersList"
$SPO_LISTS += "Eionet-Organizations-List"
#>

# Connect to the SPO site.
# User will be asked for credentials if not already authenticated.
# See https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html
Write-Output (-join("Connecting to the SPO site: ", $SPO_SITE_PATH, "..."))
# "-Interactive": Connects to the Azure AD using interactive login, allowing you to authenticate using MFA.
Connect-PnPOnline -Url $SPO_SITE_PATH -Interactive

If (-not(Get-PnPContext)) {
    Write-Output "Error connecting to SharePoint Online, unable to establish context!"
    Write-Output "Exiting..."
    Exit
}

# Output the results
Write-Output "--------------------------------------"
Foreach ($list_name in $SPO_LISTS) {
    $list_rows_count = (Get-PnPListItem -List $list_name -PageSize 5000).Count
    Write-Output (-join("`t", $list_name, " (", $list_rows_count, " rows)"))
}
