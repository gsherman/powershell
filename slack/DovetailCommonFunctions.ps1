function configure-logging()
{
	$LogManager = [FChoice.Common.LogManager]
	$LogManager::LogConfigFilePath = "C:\Dovetail\logging.config" 
	$LogManager::Reconfigure()
}

function create-clarify-session
{
	$clarifyapp = $args[0]
	$ClarifySession = $clarifyapp::Instance.CreateSession();
	$ClarifySession.TruncateStringFields = "True";
	return $ClarifySession
}

function create-clarify-application
{
	#Load the configuration file
	.\LoadConfig dovetail.config
	
	$connectionString = $appSettings["connectionString"];
	$databaseType = $appSettings["databaseType"];    	
	
	[system.reflection.assembly]::LoadWithPartialName("fcsdk") > $null;
	[system.reflection.assembly]::LoadWithPartialName("FChoice.Toolkits.Clarify") > $null;
	
	$config = new-object -typename System.Collections.Specialized.NameValueCollection;
	$config.Add("fchoice.connectionstring",$connectionString);
	$config.Add("fchoice.dbtype",$databaseType);
	$config.Add("fchoice.disableloginfromfcapp", "false"); 
	$config.Add("fchoice.nocachefile", "true"); 
	
	#Create and initialize the clarifyApplication object
	$ClarifyApplication = [Fchoice.Foundation.Clarify.ClarifyApplication];

	#if configured to do so, turn on logging
	if ($appSettings["enableLogging"] -eq $true)
	{
		write-debug "logging is enabled";
		configure-logging;
	}

	if ($ClarifyApplication::IsInitialized -eq $false ){
	   $ClarifyApplication::initialize($config) > $null;
	}
	return $ClarifyApplication;
}

function display-fcsdk-version()
{
	$assembly = [System.Reflection.Assembly]::LoadWithPartialName("fcsdk")
	$assemblyName = $assembly.GetName()
	$assemblyVersion =  $assemblyName.version
	"Assembly: {0} has version number of: {1}" -f $assemblyName.name, $assemblyVersion
}
	
function dispatch-dialogue ( $ClarifySession, [string] $dialogueId, [string] $queueName)
{
	write-debug "Dispatching dialogue $dialogueId to queue $queueName.";
	
	$dispatcherUserName = 'sa';	

	$dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet $ClarifySession;
	$workflowManager = new-object FChoice.Foundation.Clarify.Workflow.WorkflowManager $dataSet;
	$now = [FChoice.Foundation.FCGeneric]::NOW_DATE;
	$result = $workflowManager.Dispatch($dialogueId, "dialogue", $queueName, $now, $dispatcherUserName, $true);	
}

function yank-case ( $ClarifySession, [string] $caseId)
{
	write-debug "Yanking dialogue $caseId";
	
	$userName = $ClarifySession.UserName;
	$wipbinName = "";

	$dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet $ClarifySession;
	$workflowManager = new-object FChoice.Foundation.Clarify.Workflow.WorkflowManager $dataSet;
	$now = [FChoice.Foundation.FCGeneric]::NOW_DATE;
	$result = $workflowManager.Yank($caseId, "case", $wipbinName, $now, $userName, $true);	
}

function auto-dispatch-case ( $ClarifySession, [string] $caseId)
{
	write-debug "Auto-Dispatching case $caseId";
	
	$rule = [FChoice.Foundation.Clarify.AutoDest.AutoDestRule]::RetrieveRule("case","EMC_DISPATCH"); 
	$queues = $rule.EvaluateRule($caseId);
	$queue = $queues[0];

	write-debug "Dispatching case $caseId to queue $queue";
		
	$dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet $ClarifySession;
	$workflowManager = new-object FChoice.Foundation.Clarify.Workflow.WorkflowManager $dataSet;
	$now = [FChoice.Foundation.FCGeneric]::NOW_DATE;
	$dispatcherUserName = $ClarifySession.UserName;
	$result = $workflowManager.Dispatch($caseId, "case", $queue, $now, $dispatcherUserName, $true);	
}

function change-case-severity($ClarifySession, [string] $caseId, [string] $severity){
	$supportToolkit= new-object FChoice.Toolkits.Clarify.Support.SupportToolkit( $ClarifySession )
	$updateCaseSetup = new-object FChoice.Toolkits.Clarify.Support.UpdateCaseSetup($caseId)
	$updateCaseSetup.Severity = $severity;
	$result = $supportToolkit.UpdateCase($updateCaseSetup);
}


function get-case-by-id([string] $id)
{
	$dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet($ClarifySession)
	$caseGeneric = $dataSet.CreateGeneric("case")
	$caseGeneric.AppendFilter("id_number", "Equals", $id)
	$caseGeneric.Query();
	$caseGeneric.Rows[0];
}

function get-caseview-by-objid([int] $objid)
{
	$dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet($ClarifySession)
	$caseGeneric = $dataSet.CreateGeneric("qry_case_view")
	$caseGeneric.AppendFilter("elm_objid", "Equals", $objid)
	$caseGeneric.Query();
	$caseGeneric.Rows[0];
}

function get-row-for-table-by-objid([string] $tableName, [int] $objid)
{
	$dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet($ClarifySession)
	$generic = $dataSet.CreateGeneric($tableName)
	$generic.AppendFilter("objid", "Equals", $objid)
	$generic.Query();
	$generic.Rows[0];
}

function get-row-for-table-by-unique-field([string] $tableName, [string] $uniqueFieldName, $fieldValue)
{
	$dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet($ClarifySession)
	$generic = $dataSet.CreateGeneric($tableName)
	$generic.AppendFilter($uniqueFieldName, "Equals", $fieldValue)
	$generic.Query();
	$generic.Rows[0];
}

function change-case-status ($ClarifySession, [string] $caseId, [string] $newStatus)
{
	write-debug "Changing Case Status for Case $caseId";
	
	$userName = $ClarifySession.UserName;
	$chgStatusNotes = "Case has been updated by the customer"
	$dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet $ClarifySession;
	$workflowManager = new-object FChoice.Foundation.Clarify.Workflow.WorkflowManager $dataSet;
	$now = [FChoice.Foundation.FCGeneric]::NOW_DATE;
	$result = $workflowManager.ChangeStatus($caseId, "case", $newStatus, $now, $userName, $true, $chgStatusNotes);	
}