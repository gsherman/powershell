param( [string] $caseId, [string] $userName) 

#source other scripts
. .\DovetailCommonFunctions.ps1
   
$ClarifyApplication = create-clarify-application; 
$ClarifySession = create-clarify-session $ClarifyApplication; 

$supportToolkit= new-object FChoice.Toolkits.Clarify.Support.SupportToolkit( $ClarifySession )
$initialResponseSetup = new-object FChoice.Toolkits.Clarify.Support.InitialResponseSetup($caseId)
$initialResponseSetup.UserName  = $userName;
$initialResponseSetup.IsVIAPhone   = $TRUE;
$result = $supportToolkit.InitialResponse($initialResponseSetup);
    
    