param( [string] $caseId) 

#source other scripts
. .\DovetailCommonFunctions.ps1
   
$ClarifyApplication = create-clarify-application; 
$ClarifySession = create-clarify-session $ClarifyApplication; 

$supportToolkit= new-object FChoice.Toolkits.Clarify.Support.SupportToolkit( $ClarifySession )
$closeCaseSetup = new-object FChoice.Toolkits.Clarify.Support.CloseCaseSetup($caseId)
$result = $supportToolkit.CloseCase($closeCaseSetup);
    
    