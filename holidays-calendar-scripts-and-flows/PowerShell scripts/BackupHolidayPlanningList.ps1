# -------------------------------------------------------------------------
# This PS script is used to manually backup SPO lists.
# It saves the content of each SPO list in a backup .xml file.
# The script uses the PnP PowerShell module - see https://pnp.github.io/powershell/
# -------------------------------------------------------------------------

$ErrorActionPreference = "Stop"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$TENANT = "EEA1.eea.com" #(EEA's tenant)

$PNP_APP_ENTRAID = "" #(EEA's tenant)

$SPO_SITE_PATH = "https://EEA1.sharepoint.com/sites/EIONETPortal" #(EEA's tenant)

$SPO_SUMMARY_LIST = "HolidayPlanningList"

$SPO_LISTS = @()
$SPO_LISTS += $SPO_SUMMARY_LIST

# Ensure that the PnP PS module is installed.
# Write-Output "Uninstall the PnP module..."
# Uninstall-Module -Name PnP.PowerShell -Force
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

    #Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -Confirm:$False
    Install-Module -Name PnP.PowerShell -Scope CurrentUser
    Write-Output "Ensure that the PnP module functions have access to the required tenant resources..."
    Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnPLibrary" -Tenant $TENANT -Interactive
    exit
} Else {
    Write-Output "The PnP module is already installed."
    # Write-Output "Checking for PnP updates..."
    # Update-Module -Name PnP.PowerShell -Scope CurrentUser -Confirm
}

# Connect to the SPO site. User will be asked for credentials if not already authenticated.
Write-Output (-join("Connecting to the SPO site: ", $SPO_SITE_PATH, "..."))
Connect-PnPOnline -Url $SPO_SITE_PATH -Interactive -ClientId $PNP_APP_ENTRAID
If (-not(Get-PnPContext)) {
    Write-Output "Error connecting to SharePoint Online site, unable to establish context!"
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
        Write-Output (-join("Error! List ", $list_name, " not found! Please verify the site url: ", $SPO_SITE_PATH))
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
        Write-Output (-join("Error creating the list template file: ", $_.Exception.Message))
        Write-Output "Exiting..."
        Exit
    }
    
    # Save the list content to the previously created list template file
    Try {
        # Display list rows count
        $list_rows_count = (Get-PnPListItem -List $list_name -PageSize 9999).Count
        Write-Output (-join("`tExporting ", $list_rows_count, " row(s)..."))

        Add-PnPDataRowsToSiteTemplate -Path $backup_list_name -List $list_name
    } Catch {
        Write-Output (-join("Error saving list data to file: ", $_.Exception.Message))
        Write-Output "Exiting..."
        Exit
    }
}

# Output the results
Write-Output "----------------------------------------"
Write-Output "The following backup .xml files were created:"
Foreach ($list_name in $SPO_LISTS) {
    Write-Output (-join("`t", $list_name, ".xml"))
}

# Disconnects the current connection and clears its token cache.
Disconnect-PnPOnline

Write-Output "Done."
