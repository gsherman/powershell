#source other scripts
. .\DovetailCommonFunctions.ps1

$ClarifyApplication = create-clarify-application; 
$ClarifySession = create-clarify-session $ClarifyApplication; 
$now = Get-Date;

$dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet($ClarifySession)
$contractGeneric = $dataSet.CreateGeneric("contract")
$contractGeneric.AppendFilter("expire_date", "LessThan", $now)
$contractGeneric.AppendFilter("status", "Equals", "Open")
$contractGeneric.Query()
$rows = $contractGeneric.Rows;

log-info ("Found " + $rows.Count + " contracts to be updated");

 foreach ($row in $rows) {
  $row["status"] = "Closed";
  log-info ("Updating contract '" + $row["id"] + "' with status of " + $row["status"] + " and expire date of " + $row["expire_date"])
 }

$contractGeneric.UpdateAll();
    
