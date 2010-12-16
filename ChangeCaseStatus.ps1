param( [string] $caseId, [string] $newStatus) 

#source other scripts
. .\DovetailCommonFunctions.ps1
   
$ClarifyApplication = create-clarify-application; 
$ClarifySession = create-clarify-session $ClarifyApplication; 

$supportToolkit= new-object FChoice.Toolkits.Clarify.Support.SupportToolkit( $ClarifySession )
$changeCaseStatusSetup = new-object FChoice.Toolkits.Clarify.Support.ChangeCaseStatusSetup ($caseId)
$changeCaseStatusSetup.NewStatus = $newStatus
$result = $supportToolkit.ChangeCaseStatus($changeCaseStatusSetup );
    
