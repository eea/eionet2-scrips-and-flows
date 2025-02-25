# -------------------------------------------------------------------------
# This PS script is used to manually restore SPO lists.
# It restores the content of each SPO list from the backup .xml file.
# The script uses the PnP PowerShell module - see https://pnp.github.io/powershell/
# -------------------------------------------------------------------------

$ErrorActionPreference = "Stop"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$TENANT = "EEA1.eea.com" #(EEA's tenant)

$PNP_APP_ENTRAID = "" #(EEA's tenant)

$SPO_SITE_PATH = "https://EEA1.sharepoint.com/sites/EIONETPortal" #(EEA's tenant)

$SPO_HOLIDAYS_LIST = "NationalHolidaysList"
$SPO_SUMMARY_LIST = "HolidayPlanningList"

$SPO_LISTS = @()
$SPO_LISTS += $SPO_HOLIDAYS_LIST
#$SPO_LISTS += $SPO_SUMMARY_LIST

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

$restored_lists = @()
# Process each predefined list
Foreach ($list_name in $SPO_LISTS) {
    Write-Output (-join("Restoring list ", $list_name, "..."))

    $backup_list_name = -join($list_name, ".xml")

    # Check if the list backup file exists
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
        Write-Output (-join("`Restoring list from file ", $backup_list_name, "..."))
        Invoke-PnPSiteTemplate -Path $backup_list_name -Handlers Lists
        $restored_lists += $list_name
    } Catch {
        #Write-Output $_
        Write-Output (-join("`tError restoring list ", $list_name))
        Write-Output (-join("`t", $_))
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
        $list_rows_count = (Get-PnPListItem -List $list_name -PageSize 5000).Count
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
Disconnect-PnPOnline

Write-Output "Done."
