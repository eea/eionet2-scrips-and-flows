$ErrorActionPreference = "Stop"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$TENANT = "EEA1.eea.com" #(EEA's tenant)

$PNP_APP_ENTRAID = "" #(EEA's tenant)

$SPO_SITE_PATH = "https://EEA1.sharepoint.com/sites/EIONETPortal" #(EEA's tenant)

$SPO_HOLIDAYS_LIST = "NationalHolidaysList"

$SPO_HOLIDAYS_LIST_IN_URL = "NationalHolidaysList" #(EEA's tenant)
$SPO_SUMMARY_LIST = "HolidayPlanningList"

# Define NationalHolidaysList internal columns names
$SPO_HOLIDAYS_LIST_COLUMN_ID = "ID"
$SPO_HOLIDAYS_LIST_COLUMN_COUNTRY = "field_0"
$SPO_HOLIDAYS_LIST_COLUMN_DATE = "field_1"
$SPO_HOLIDAYS_LIST_COLUMN_NAME = "Title"
$SPO_HOLIDAYS_LIST_COLUMN_LOCAL_NAME = "field_5"

# Define HolidayPlanningList internal columns names
$SPO_SUMMARY_LIST_COLUMN_TITLE = "Title"
$SPO_SUMMARY_LIST_COLUMN_START_DATE = "field_3"
$SPO_SUMMARY_LIST_COLUMN_END_DATE = "field_4"
$SPO_SUMMARY_LIST_COLUMN_HOLIDAYS_LINK = "field_5"
$SPO_SUMMARY_LIST_COLUMN_COUNTRIES = "field_6"
$SPO_SUMMARY_LIST_COLUMN_EVENT_CATEGORY = "field_7"
$SPO_SUMMARY_LIST_COLUMN_EVENT_TYPE = "field_8"
$SPO_SUMMARY_LIST_COLUMN_GROUPS = "field_9"
$SPO_SUMMARY_LIST_COLUMN_ORIGINAL_EVENT_ID = "field_2"
$SPO_SUMMARY_LIST_COLUMN_SOURCE_LIST = "Source_x0020_list"

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

# Check if lists exists
Write-Output "Check if lists exist..."

$list_holidays = Get-PnPList -Identity $SPO_HOLIDAYS_LIST -ErrorAction SilentlyContinue
If ($list_holidays -eq $Null) {
    # If we can't find one of the lists we will terminate the script execution
    Write-Output (-join("Error! List ", $SPO_HOLIDAYS_LIST, " not found! Please verify the site url: ", $SPO_SITE_PATH))
    Write-Output "Exiting..."
    Exit
}
$list_summary = Get-PnPList -Identity $SPO_SUMMARY_LIST -ErrorAction SilentlyContinue
If ($list_summary -eq $Null) {
    # If we can't find one of the lists we will terminate the script execution
    Write-Output (-join("Error! List ", $SPO_SUMMARY_LIST, " not found! Please verify the site url: ", $SPO_SITE_PATH))
    Write-Output "Exiting..."
    Exit
}

Write-Output "Loading Holidays in the HolidayPlanningList list..."

Write-Output "Reading Holidays ordered by Date ASC..."
$holidays_ordered_query = [string]::Format("
    <View Scope='RecursiveAll'>
        <Query>
            <OrderBy>
                <FieldRef Name='{0}' Ascending='True' />
            </OrderBy>
        </Query>
    </View>",
    $SPO_HOLIDAYS_LIST_COLUMN_DATE)
#Write-Output $holidays_ordered_query
$holidays_list_items = Get-PnPListItem -List $list_holidays -PageSize 5000 -Query $holidays_ordered_query
Write-host (-join("Total number of Holidays found: ", $holidays_list_items.count))


$previous_holiday_date = $ZeroDate
$previous_holiday_title = ""
$countries_list = ""
$countries_count = 0
$first_item = 1 #set first item position flag

$idx = 0

$count = 0
foreach($list_item in $holidays_list_items)
{  

    $Title = $list_item[$SPO_HOLIDAYS_LIST_COLUMN_NAME]
    $Country = $list_item[$SPO_HOLIDAYS_LIST_COLUMN_COUNTRY]

    $Date = $list_item[$SPO_HOLIDAYS_LIST_COLUMN_DATE]
    $Date = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($Date, 'Romance Standard Time')

    #Write-Host $Date -ForegroundColor blue

    if ($first_item -eq 1) {
        #Nothing to do but to save this first item data to history and move to the next item
        $previous_holiday_title = $Title
        $countries_list = $Country #initialize countries list that are observing the holiday on the same date with the first item country
        $previous_holiday_date = $Date

        $countries_count = 1
        $first_item = 0 #reset first item position flag
    } else {
        #Write-Output (-join("`tLast date:    ", $previous_holiday_date.Date))
        #Write-Output (-join("`tCurrent date: ", $Date.Date))

        if($Date.Date -ne $previous_holiday_date.Date) {
            #Write-Host "`tWe moved to a different date, so we create a summary item with the saved agregated history data." -ForegroundColor red

            #Write-Output (-join("`tCountries count:", $countries_count))
            #Write-Output (-join("`tCountries list: ", $countries_list))

            $summary_title = ""
            if($countries_count -eq 1) {
                #We have a single country
                $summary_title = [string]::Format("{0} - {1}",
                                    $countries_list,
                                    $previous_holiday_title)
            } else {
                 #We have multiple countries
                $summary_title = [string]::Format("{0} countries - {1}",
                                    $countries_count,
                                    $previous_holiday_title)
            }
            #Write-Output (-join("`tSummary title:", $summary_title))

            $holidays_link = [string]::Format("{0}/Lists/{1}/AllItems.aspx?{2}{3}",
                                $SPO_SITE_PATH,
                                $SPO_HOLIDAYS_LIST_IN_URL,
                                "FilterField1=field_1&FilterType1=DateTime&FilterValue1=",
                                $previous_holiday_date.ToString('yyyy-MM-dd'))


            $holiday_start = $previous_holiday_date
            $holiday_end = $holiday_start

            
            Add-PnPListItem -List $SPO_SUMMARY_LIST -Values @{
                $SPO_SUMMARY_LIST_COLUMN_TITLE = $summary_title;
                $SPO_SUMMARY_LIST_COLUMN_START_DATE = $holiday_start;
                $SPO_SUMMARY_LIST_COLUMN_END_DATE = $holiday_end;
                $SPO_SUMMARY_LIST_COLUMN_HOLIDAYS_LINK = $holidays_link;
                $SPO_SUMMARY_LIST_COLUMN_COUNTRIES = $countries_list;
                $SPO_SUMMARY_LIST_COLUMN_EVENT_CATEGORY = "Holiday";
                $SPO_SUMMARY_LIST_COLUMN_EVENT_TYPE = "";
                $SPO_SUMMARY_LIST_COLUMN_GROUPS = "";
                $SPO_SUMMARY_LIST_COLUMN_ORIGINAL_EVENT_ID = "";
                $SPO_SUMMARY_LIST_COLUMN_SOURCE_LIST = "NationalHolidaysList";
            }
            

            $count += 1

            if($countries_count -eq 1) {
                Write-Output (-join("Created Holiday summary item - for Single country: ", $countries_list, ", ", $holiday_start))
            } else {
                Write-Output (-join("Created Holiday summary item - for Multiple countries: ", $countries_list.replace(';#', ','), ", ", $holiday_start))
            }

            #Save current data to history
            $previous_holiday_date = $Date
            $previous_holiday_title = $Title
            $countries_list = $Country
            $countries_count = 1
        } else {
            #We are on the same date as the previous holiday; Save current data to history
            #Write-Host "`tWe are on the same date - we need to aggregate countries" -ForegroundColor yellow

            #$previous_holiday_date = $Date
            #$previous_holiday_title = $Title

            $countries_list  += -join(';#', $Country)
            #Need a fix for when on the same day and in the same country there are holidays with different names
            #The fix is to make sure the country is added only once in the list
            $split_countries = $countries_list -split ';#'
            $unique_countries_list = ($split_countries | Sort-Object | Get-Unique) -join ';#'
            if($countries_list.Length -eq $unique_countries_list.Length) {
                $countries_count += 1
            } else {
                Write-Host (-join("`tCountry is already added in countries list - skipping: ", $Country)) -ForegroundColor yellow
                $countries_list = $unique_countries_list
            }

            #Write-Output (-join("`tCurrent countries list: ", $countries_list))
            #Write-Output (-join("`tCurrent countries count: ", $countries_count))
        }
    }

    $idx += 1
}

Write-Output (-join("Total number of created Holidays in the Events summary list: ", $count))

# Disconnect the current connection and clears its token cache.
# It will require you to build up a new connection again using Connect-PnPOnline in order to further use PnP PowerShell cmdlets.
Disconnect-PnPOnline

Write-Output "Done."
