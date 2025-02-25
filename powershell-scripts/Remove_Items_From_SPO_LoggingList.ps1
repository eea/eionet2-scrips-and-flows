# Set the variables

#Update-Module -Name PnP.PowerShell

$SiteURL = "https://eea1.sharepoint.com/sites/-EXT-EionetConfiguration/"
#$SiteURL = "https://2xkk2b.sharepoint.com/sites/Testsite" # Daniel's test site
	
# Connect to SharePoint Online site using modern authentication and MFA
Connect-PnPOnline -Url $SiteURL -Interactive

#$ListName = "Logging List"
$ListName = "LoggingList"
$ColumnName = "ApplicationName"
$ColumnValue = "Eionet2-User-Management" # Delete based on this category
$DateColumn = "Timestamp"
$DateValue = "2018-04-20T00:00:00" # Remove everything older than this timestamp: TIMESTAMP needs to be in this format
$BatchSize = 5000

$startTime = (Get-Date)

# Get the list
$List = Get-PnPList -Identity $ListName


#------------------ Part 1 Get the list items


$ListItems = (Get-PnPListItem -List $List -PageSize $BatchSize)
$list_rows_count = $ListItems.Count
Write-Host "Total number of items in the list: $list_rows_count"


#------------------ Part 2 Get only the list items which meet the conditions


# Set up a variable to hold all filtered items in the list
$ListFilteredItems = @()
$ListFilteredItemsByAppName  = $ListItems | Where { $_["$ColumnName"] -Eq $ColumnValue }
Write-Host "`tTotal number of items matching [ApplicationName] condition: $($ListFilteredItemsByAppName.Count)"
$ListFilteredItems = $ListFilteredItemsByAppName | Where { $_["$DateColumn"] -Lt $DateValue }
Write-Host "`tTotal number of items matching both [ApplicationName] and [Timestamp] conditions: $($ListFilteredItems.Count)"

# Print out the total number of items to delete
Write-Host "Total number of items to delete: $($ListFilteredItems.Count)"


#------------------ Part 3 Delete the filtered list items


# Ask for confirmation before deleting the items
$Confirmation = Read-Host "Are you sure you want to delete the items? (y/n)"
if ($Confirmation -eq "y") {
    $Batch = New-PnPBatch

    # Clear items in the List
    ForEach ($Item in $ListFilteredItems)
    {   
        #Write-Host "$($Item.ID) (Timestamp: $($Item["Timestamp"]))"
        Remove-PnPListItem -List $ListName -Identity $Item.ID -Batch $Batch
    }
 
    # Send Batch to the server
    Write-Host "Sent Batch delete to the server..."
    Invoke-PnPBatch -Batch $Batch
    Write-Host "Done."
}


$ListItems = (Get-PnPListItem -List $List -PageSize $BatchSize)
$list_rows_count = $ListItems.Count
# Print out the new total number of items in the list
Write-Host "Total number of items in the list after delete: $list_rows_count"

$endTime = (Get-Date)
Write-Host (-join("Total script runtime: ", ($endTime-$startTime).ToString('''''mm'' min. ''ss'' sec.''')))

