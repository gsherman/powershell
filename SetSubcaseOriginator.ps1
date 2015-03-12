##############################################################################
## SCRIPT:	SetSubcaseOriginator.ps1
##
## PURPOSE:	SetSubcaseOriginator - workaround for a task manager bug
##
##############################################################################
param( [string] $id, [string] $originatorLoginName) 

#source other scripts
. .\DovetailCommonFunctions.ps1

#validate input parameters
if ($id -eq "" -or $originatorLoginName -eq "") 
{
   write-host "Error: Missing Required parameters."
   write-host "Usage:"
   write-host
   write-host "SetSubcaseOriginator.ps1 subcaseId originatorLoginName"
   write-host
   exit
}

function get-subcase-by-id([string] $id)
{
  $dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet($ClarifySession)
  $subcaseGeneric = $dataSet.CreateGeneric("subcase")
  $subcaseGeneric.AppendFilter("id_number", "Equals", $id)
  $subcaseGeneric.Query();
  $subcaseGeneric.Rows[0];
}

function get-user-by-loginName([string] $loginName)
{
  $dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet($ClarifySession)
  $userGeneric = $dataSet.CreateGeneric("user")
  $userGeneric.AppendFilter("login_name", "Equals", $loginName)
  $userGeneric.Query();
  $userGeneric.Rows[0];
}
   
$ClarifyApplication = create-clarify-application; 
$ClarifySession = create-clarify-session $ClarifyApplication; 

$subcase = get-subcase-by-id($id);

if ($subcase -eq $null)
{
   write-host "Error: Could not locate subcase with id of $id" 
   exit
}

$user = get-user-by-loginName($originatorLoginName);
if ($user -eq $null)
{
   write-host "Error: Could not locate user with loginName of $originatorLoginName"; 
   exit
}

$userObjid = $user["objid"];
$subcase["subc_orig2user"] = $userObjid;
$subcase.Update();
