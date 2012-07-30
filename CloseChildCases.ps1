$connectionString = "Data Source=localhost;Initial Catalog=dovetail;uid=sa;pwd=sa"
$databaseType = "MSSQL"
 
[system.reflection.assembly]::LoadWithPartialName("fcsdk") > $null
[system.reflection.assembly]::LoadWithPartialName("FChoice.Toolkits.Clarify") > $null

$config = new-object -typename System.Collections.Specialized.NameValueCollection
$config.Add("fchoice.connectionstring",$connectionString);
$config.Add("fchoice.dbtype",$databaseType);
$config.Add("fchoice.disableloginfromfcapp", "false"); 

$ClarifyApplication = [Fchoice.Foundation.Clarify.ClarifyApplication]

if ($ClarifyApplication::IsInitialized -eq $false ){
   $ClarifyApplication::initialize($config) > $null;
}

$ClarifySession = $ClarifyApplication::Instance.CreateSession()
$supportToolkit= new-object FChoice.Toolkits.Clarify.Support.SupportToolkit( $ClarifySession )   

$parentCaseId = [string] $args[0]

# Get the parent case (supercase)
   $dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet($ClarifySession)
   $parentCaseGeneric = $dataSet.CreateGeneric("case")
   $parentCaseGeneric.AppendFilter("id_number", "Equals", $parentCaseId)
   $parentCaseGeneric.Query()
   if ($parentCaseGeneric.Rows.Count -eq 0){
      write-host
			write-host "Case ID $parentCaseId not found."  -backgroundcolor "red"
			write-host
     	exit
   }
   $parentCaseObjid = $parentCaseGeneric.Rows[0]["objid"]
   
# Query for open children cases for the given parent case
	 $childCaseGeneric = $dataSet.CreateGeneric("victimcase")
   $childCaseGeneric.AppendFilter("supercase_objid", "Equals", $parentCaseObjid);
   $childCaseGeneric.AppendFilter("condition", "Like", "Open"); 
   $childCaseGeneric.Query()

	 write-host Found $childCaseGeneric.Rows.Count open child cases for parent case $parentCaseId 		   
   
# For each child case, close it
		 foreach( $childCase in $childCaseGeneric.Rows){
		   write-host Closing Case $childCase["id_number"]		   
       $closeCaseResult = $supportToolkit.CloseCase($childCase["id_number"]);
		 }
