# -------------------------------------------------------------------------
# This PS script is used to backup the SPO lists used by the Eionet 2 apps.
# It saves the content of each SPO list in a backup .xml file.
# At the end, it creates a zip archive with all the backup xml files.
# The script uses the PnP PowerShell module - see https://pnp.github.io/powershell/
# Last update: April 2023
# -------------------------------------------------------------------------

# IMPORTANT! YOU MUST UPDATE THIS WITH YOUR SPO SITE URL (where the Eionet 2 lists are stored)
$SPO_SITE_PATH = "https://2xkk2b.sharepoint.com/sites/Testsite" # (Daniel's tenant)
#$SPO_SITE_PATH = "https://7lcpdm.sharepoint.com/sites/EIONETPortal" # (Mihai's tenant)
#$SPO_SITE_PATH = "https://<...>.sharepoint.com" # (EEA's tenant)

$SPO_LISTS = @()
# IMPORTANT! If you need to add/remove a list, make sure you use its "Display name" property (not its "Name" property).
<#
$SPO_LISTS += "ConfigurationList"
$SPO_LISTS += "ConsultationsList"
$SPO_LISTS += "Events Participants"
$SPO_LISTS += "UsersList"
$SPO_LISTS += "MappingList"
$SPO_LISTS += "EventsList"
#$SPO_LISTS += "LoggingList" #This is a huge list. During testing you might want to skip it by commenting this line.
$SPO_LISTS += "OrganisationList"
#>
# Some local test lists - to be removed in the script release version
$SPO_LISTS += "Eionet-Organizations-List"
$SPO_LISTS += "EionetUsersList"

# Ensure that the PnP PS module is installed in your system.
# See https://pnp.github.io/powershell/
#Write-Output "Uninstall the PnP module..."
#Uninstall-Module -Name PnP.PowerShell -Force
Write-Output "Ensuring that the PnP module is installed..."
If (-not(Get-Module -Name PnP.PowerShell -ListAvailable)) {
    Write-Output "The PnP module was not found - installing..."

    # We need to install the PnP module
    
    If ($IsWindows) {
        # Check first if script is executed under elevated permissions - Run as Administrator
        If (-not([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
            Write-Output "Please run this script in elevated mode (Run as Administrator)!"
            Write-Output "Exiting..."
            Exit
        }
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

$backup_files_list = @()
# Process each list
Foreach ($list_name in $SPO_LISTS) {
    Write-Output (-join("Processing list ", $list_name, "..."))

    # Check if list exists
    $list = Get-PnPList -Identity $list_name -ErrorAction SilentlyContinue
    If ($List -eq $Null) {
        # If we can't find one of the lists we will terminate the script execution
        Write-Output (-join("`tError! List ", $list_name, " not found! Please verify the site url: ", $SPO_SITE_PATH))
        Write-Output "Exiting..."
        Exit
    }

    $backup_list_name = -join($list_name, ".xml")
    $backup_files_list += ,$backup_list_name

    # Delete list backup file if exists
    If (Test-Path $backup_list_name) {
        Write-Output (-join("`tRemoving existing file ", $backup_list_name, "..."))
        Remove-Item $backup_list_name
    }

    # Create list template file (with the list structure definition only)
    Try {
        Get-PnPSiteTemplate -Out $backup_list_name -ListsToExtract $list_name -Handlers Lists
    } Catch {
        Write-Output (-join("`tError creating the list template file: ", $_.Exception.Message))
        Write-Output "Exiting..."
        Exit
    }
    
    # Save the list content to the previously created list template file
    Try {
        # Display list rows count
        $list_rows_count = (Get-PnPListItem -List $list_name).Count
        Write-Output (-join("`tExporting ", $list_rows_count, " row(s)..."))

        Add-PnPDataRowsToSiteTemplate -Path $backup_list_name -List $list_name
    } Catch {
        Write-Output (-join("`tError saving the list items to file: ", $_.Exception.Message))
        Write-Output "Exiting..."
        Exit
    }
}

# Create a zip archive with the lists backup files
Write-Output "Creating the zip archive..."
#$timestamp = (Get-Date).ToString().Replace(",","_").Replace(" ","_").Replace(":","-")
$timestamp = ((Get-Date -UFormat %s) -replace("[^0-9]")).ToString()
$compress_archive_name = -join("SPO_Lists_Backup_", $timestamp, ".zip")
Foreach ($filename in $backup_files_list)
{
    If (Test-Path $filename) {
        Try {
            Write-Output (-join("`tAdding file ", $filename, " to ", $compress_archive_name, "..."))
            Compress-Archive -Update $filename $compress_archive_name
        } Catch {
            Write-Output (-join("Error adding file ", $filename, " to the zip archive!"))
            Write-Output $_.Exception.Message
            Write-Output "Exiting..."
            Exit
        }
    } Else {
        Write-Output (-join("Error! File not found: ", $filename))
        Write-Output "Exiting..."
        Exit
    }
}

# Output the results
Write-Output "----------------------------------------"
Write-Output "The following backup files were created:"
Foreach ($list_name in $SPO_LISTS) {
    Write-Output (-join("`t", $list_name, ".xml"))
}
If (Test-Path $compress_archive_name) {
    Write-Output (-join("The lists backup files were added to this archive: ", (Get-Location).ToString(), "\", $compress_archive_name))
} Else {
    Write-Output "Error! The zip archive file was not created."
}

Write-Output "`nIMPORTANT! Please copy the backup file(s) to a safe backup location!"
Write-Output "[And please verify first the content of the zip archive.]"

# Disconnects the current connection and clears its token cache.
# It will require you to build up a new connection again using Connect-PnPOnline in order to use any of the PnP PowerShell cmdlets.
# You will have to reauthenticate.
#Disconnect-PnPOnline
