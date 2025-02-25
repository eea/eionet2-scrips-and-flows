$ErrorActionPreference = "Stop"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$TENANT = "EEA1.eea.com" #(EEA's tenant)

$PNP_APP_ENTRAID = "" #(EEA's tenant)

$SPO_SITE_PATH = "https://EEA1.sharepoint.com/sites/EIONETPortal" #(EEA's tenant)

# Define NationalHolidaysList internal columns names
$SPO_HOLIDAYS_LIST_COLUMN_ID = "ID"
$SPO_HOLIDAYS_LIST_COLUMN_COUNTRY = "field_0"
$SPO_HOLIDAYS_LIST_COLUMN_DATE = "field_1"
$SPO_HOLIDAYS_LIST_COLUMN_NAME = "Title"
$SPO_HOLIDAYS_LIST_COLUMN_LOCAL_NAME = "field_5"


$SPO_HOLIDAYS_LIST = "NationalHolidaysList"

$csv_file = ".\eu-national-holidays.csv"

$countries = @('AL','AT','BA','BE','BG','CY','CZ','DE','DK','EE','GR','ES','FI','FR','GB','HR','HU','IE','IS','IT','LI','LT','LU','LV','ME','MK','MT','NL','PL','PT','RO','RS','SE','SI','SK','TR','XK')

$years = @(2024,2025,2026)

$total_holidays = 0
$total_local_holidays = 0

if (Test-Path $csv_file) {
    Remove-Item $csv_file -verbose
}


for ($country_idx = 0; $country_idx -lt $countries.Length; $country_idx++) {
    if($countries[$country_idx] -eq 'XK') {
        # Kosovo - No data available
        continue
    }

    for ($year_idx = 0; $year_idx -lt $years.Length; $year_idx++) {
        $url = "https://date.nager.at/api/v3/publicholidays/" + $years[$year_idx] + "/" + $countries[$country_idx]
        #Write-Host $url
        $Holidays = Invoke-RestMethod -Method Get -UseBasicParsing -Uri $url

        for ($i = 0; $i -lt $holidays.Length; $i++) {
            if($holidays[$i].global -ne 'true') {
                Write-Host "  Skip:", $holidays[$i].countryCode, ",", $holidays[$i].name, ",", $holidays[$i].localName
                $total_local_holidays += 1
                continue
            }


            if($holidays[$i].countryCode -eq 'GR') {
                $hol_country = 'EL'
            } else {
                $hol_country = $holidays[$i].countryCode
            }
            $hol_date =  [datetime]::parseexact($holidays[$i].date, 'yyyy-MM-dd', $null)
            $hol_date_str = $hol_date.ToString("yyyy/MM/dd")
            $hol_year_str = $hol_date.ToString("yyyy")
            $hol_name = $holidays[$i].name
            $hol_local_name = $holidays[$i].localName

            Write-Host "Add:", $holidays[$i].countryCode, ",", $hol_date_str, ",", $hol_year_str, ",", $hol_name, ',', $hol_local_name
            $total_holidays += 1

            [PsCustomObject]@{
                Country = $holidays[$i].countryCode
                Date = $hol_date
                Year = [Int]$hol_year_str
                Name = $hol_name
                LocalName = $hol_local_name
            } | Export-Csv -Append -Encoding UTF8BOM -NoTypeInformation -Path $csv_file
        }
        Write-Host "Write To CSV - ", $countries[$country_idx], ":", $years[$year_idx], ":", $holidays.Length
    }
}

Write-Host
Write-Host "Total local holidays skipped:", $total_local_holidays
Write-Host "Total holidays found:", $total_holidays


# Ensure that the PnP PS module is installed.
# See https://pnp.github.io/powershell/
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


# Connect to the SPO site.
# User will be asked for credentials if not already authenticated.
# See https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html
Write-Output (-join("Connecting to the SPO site: ", $SPO_SITE_PATH, "..."))
# "-Interactive": Connects to the Azure AD using interactive login, allowing you to authenticate using MFA.
Connect-PnPOnline -Url $SPO_SITE_PATH -Interactive -ClientId $PNP_APP_ENTRAID

If (-not(Get-PnPContext)) {
    Write-Output "Error connecting to SharePoint Online site, unable to establish context!"
    Write-Output "Exiting..."
    Exit
}

Write-Output (-join("Preparing holidays list ", $SPO_HOLIDAYS_LIST, "..."))

# Check if list exists
$list = Get-PnPList -Identity $SPO_HOLIDAYS_LIST -ErrorAction SilentlyContinue
If ($List -eq $Null) {
    # If we can't find one of the required lists we will terminate the script execution
    Write-Output (-join("Error! List ", $SPO_HOLIDAYS_LIST, " not found! Please verify the site url: ", $SPO_SITE_PATH))
    Write-Output "Exiting..."
    Exit
}

# Get all existing items/holidays in the list
$items = Get-PnPListItem -List $SPO_HOLIDAYS_LIST -PageSize 9999

# Delete any existing items
Write-Output "Deleting existing holidays in the list..."
# Create a (delete all items) batch
$Batch = New-PnPBatch
# Delete alll items in the list
ForEach($Item in $items)
{    
    Remove-PnPListItem -List $SPO_HOLIDAYS_LIST -Identity $Item.ID -Recycle -Batch $Batch
}
# Send batch to the server and execute it
Invoke-PnPBatch -Batch $Batch
Write-Output (-join("Deleted ", $items.Length, " item(s)..."))

$items_count = 0

Write-Output (-join("Importing holidays from ", $csv_file, "..."))
Import-Csv -Path $csv_file | 
    ForEach-Object {
        # Fix EL country code
        if($_.Country -eq "GR") {
            $Country = "EL"
        } else {
            $Country = $_.Country
        }
        $Date = ([datetime]$_.Date).Date
        $Name = $_.Name
        $LocalName = $_.LocalName

        Write-Host "CSV - ", $Date, ":", $Country, ":", $Name, ":", $LocalName

        Add-PnPListItem -List $SPO_HOLIDAYS_LIST -Values @{
            $SPO_HOLIDAYS_LIST_COLUMN_COUNTRY = $Country;
            $SPO_HOLIDAYS_LIST_COLUMN_DATE = $Date;
            $SPO_HOLIDAYS_LIST_COLUMN_NAME = $Name;
            $SPO_HOLIDAYS_LIST_COLUMN_LOCAL_NAME = $LocalName
        }
        $items_count += 1
    }

Write-Output (-join("Imported ", $items_count, " holidays(s) in the list."))

# Disconnect the current connection and clears its token cache.
# It will require you to build up a new connection again using Connect-PnPOnline in order to further use PnP PowerShell cmdlets.
Disconnect-PnPOnline

Write-Output "Done."
