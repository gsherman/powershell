##############################################################################
## SCRIPT:	UpdateLastModifyTimestamp.ps1
##
## PURPOSE:	This PowerShell script updates the last modified timestamp of an 
##					entity to the current date/time
##
## Call this script like: 
##		UpdateLastModifyTimestamp.ps1 objectType objid 
##
##############################################################################
param( [string] $tableName,  [int] $objid) 

#source other scripts
. .\DovetailCommonFunctions.ps1

function get-lastModifyTimestampFieldName-for-table([string] $tableName)
{
switch ($tableName) 
    { 
        "case" 				{"modify_stmp"} 
        "subcase" 		{"modify_stmp"} 
        "bug" 				{"modify_stmp"} 
        "demand_hdr" 	{"modify_stmp"} 
        "bus_org" 		{"update_stamp"} 
        "contact" 		{"update_stamp"} 
        "site" 				{"update_stamp"} 
        "task" 				{"update_stamp"} 
        "template" 		{"update_stamp"}         
        "contract" 		{"last_update"} 
        "lead" 				{"last_update"} 
        default {""}
    }
}

#validate input parameters
if ($objid -eq "" -or $tableName -eq "") 
{
   write-host "Error: Missing Required parameters."
   write-host "Usage:"
   write-host
   write-host "UpdateLastModifyTimestamp.ps1 objectType objid"
   write-host
   exit
}
   
$ClarifyApplication = create-clarify-application; 
$ClarifySession = create-clarify-session $ClarifyApplication; 

$lastModifyTimestampFieldName = get-lastModifyTimestampFieldName-for-table($tableName);
if ($lastModifyTimestampFieldName -eq "")
{
   write-host "Error: Could not determine the last Modify Timestamp Field Name for table named $tableName"
   exit
}

$row = get-row-for-table-by-objid $tableName $objid;
if ($row -eq $null)
{
   write-host "Error: Could not locate $tableName with objid of $objid" 
   exit
}

$row[$lastModifyTimestampFieldName] = Get-date;
$row.Update();
    