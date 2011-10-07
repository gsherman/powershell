param( [int] $caseObjid, [string] $action) 

#source other scripts
. .\DovetailCommonFunctions.ps1
   
$ClarifyApplication = create-clarify-application; 
$ClarifySession = create-clarify-session $ClarifyApplication; 

$dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet($ClarifySession)
$caseClockGeneric = $dataSet.CreateGeneric("case_clock")
$caseClockGeneric.AppendFilter("case_clock2case", "Equals", $caseObjid)
$caseClockGeneric.Query();

$isRunning = $false;

#if there's not yet a case_clock row for this case, add a new row
if ($caseClockGeneric.Rows.Count -eq 0){
	$caseClockRow = $caseClockGeneric.AddNew();
	$caseClockRow["time_so_far"] = 0;
	$caseClockRow["case_clock2case"] = $caseObjid;
}
else{
	$caseClockRow = $caseClockGeneric.Rows[0];
	if ($caseClockRow["is_running"] -eq 1){
		$isRunning = $true;
	}
	
}

if ($action -eq "start"){

	#start the clock

	if ($isRunning -eq $true){exit;}
	$caseClockRow["is_running"] = 1;
	$caseClockRow["status"] = "running";
	$caseClockRow["last_started_at"] = Get-date;
}
else{

	#stop/pause the clock

	if ($isRunning -eq $false){exit;}
	
	$caseClockRow["is_running"] = 0;
	$caseClockRow["status"] = "paused";
	
	$diff = new-TimeSpan $caseClockRow["last_started_at"] $(Get-Date);
	$caseClockRow["time_so_far"]+= [int]$diff.TotalSeconds;
	 
	$caseClockRow["last_started_at"] = '1/1/1753';
	$caseClockRow["last_calculated_time"] = Get-date;		
}

$caseClockRow.Update();

    