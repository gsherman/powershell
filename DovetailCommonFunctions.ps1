function configure-logging()
{
	$LogManager = [FChoice.Common.LogManager]
	$LogManager::LogConfigFilePath = $appSettings["logConfigFilePath"];
	$LogManager::Reconfigure()
	$global:logger = $LogManager::GetLogger("PowerShell"); 
}

function log-info ([string] $message) 
{ 
    write-host "INFO: $message" 
    $global:logger.LogInfo($message); 
}

function log-warn ([string] $message) 
{ 
    write-warning $message
    $global:logger.LogWarn($message); 
}

function log-debug ([string] $message) 
{ 
    write-debug $message
    $global:logger.LogDebug($message); 
}

function log-error ([string] $message) 
{ 
    write-host "ERROR: $message" -foregroundcolor red 
    $global:logger.LogError($message); 
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
	[system.reflection.assembly]::LoadWithPartialName("FChoice.Common") > $null;
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

function get-caseview-by-id([string] $id)
{
	$dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet($ClarifySession)
	$caseGeneric = $dataSet.CreateGeneric("qry_case_view")
	$caseGeneric.AppendFilter("id_number", "Equals", $id)
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

function get-empl-view-by-login-name([string] $loginName)
{
	$dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet($ClarifySession)
	$generic = $dataSet.CreateGeneric("empl_view")
	$generic.AppendFilter("login_name", "Equals", $loginName)
	$generic.Query();
	$generic.Rows[0];
}
