#Parameters
$SiteURL= "..."
$ListName = "..."
$TemplateFileName = "ListTemplate.stp"
$TemplateName = "..."
$TemplateDescription="..."
$IncludeData = $True
  
#Connect to PnP Online
Connect-PnPOnline -Url $SiteURL -Interactive
$Context  = Get-PnPContext
 
#Get the List
$List = Get-PnpList -Identity $ListName
 
#Save List as template
$List.SaveAsTemplate($TemplateFileName, $TemplateName, $TemplateDescription, $IncludeData)
$Context.ExecuteQuery()
