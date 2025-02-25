$ErrorActionPreference = "Stop"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$TENANT = "EEA1.eea.com" #(EEA's tenant)

$PNP_APP_ENTRAID = "" #(EEA's tenant)

$SPO_SITE_PATH = "https://EEA1.sharepoint.com/sites/EIONETPortal" #(EEA's tenant)

$SPO_SUMMARY_LIST = "HolidayPlanningList"

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
#Connect-PnPOnline -Url $SPO_SITE_PATH -Interactive #Old style login
Connect-PnPOnline -Url $SPO_SITE_PATH -Interactive -ClientId $PNP_APP_ENTRAID
If (-not(Get-PnPContext)) {
    Write-Output "Error connecting to SharePoint Online site, unable to establish context!"
    Write-Output "Exiting..."
    Exit
}

Write-Output "Check if list exists..."
$list_summary = Get-PnPList -Identity $SPO_SUMMARY_LIST -ErrorAction SilentlyContinue
If ($list_summary -eq $Null) {
    Write-Output (-join("Error! List ", $SPO_SUMMARY_LIST, " not found! Please verify the site url: ", $SPO_SITE_PATH))
    Write-Output "Exiting..."
    Exit
}

Write-Output (-join("Delete all data in the Events summary list: ", $SPO_SUMMARY_LIST, "..."))
# Get all existing items in the Events summary list
$items = Get-PnPListItem -List $SPO_SUMMARY_LIST -PageSize 9999
# Loop through list items and delete each item
Write-Output "Deleting existing items in the Events summary list..."
# Create a (delete all items) batch
$Batch = New-PnPBatch
# Delete alll items in the list
ForEach($Item in $items)
{    
    Remove-PnPListItem -List $SPO_SUMMARY_LIST -Identity $Item.ID -Recycle -Batch $Batch
}
# Send batch to the server and execute it
Invoke-PnPBatch -Batch $Batch
Write-Output (-join("Deleted ", $items.Length, " item(s) from the Events summary list..."))

# Disconnect the current connection and clears its token cache.
Disconnect-PnPOnline

Write-Output "Done."
