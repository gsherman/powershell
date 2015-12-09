##############################################################################
## SCRIPT:	SolutionMessage.ps1
##
## PURPOSE:	This PowerShell script posts a solution message to slack
##
## Call this script like: 
##		SolutionMessage.ps1 solutionId channel(without the leading pound sign) message
##
##############################################################################

param([string] $solutionId, [string] $channel, [string] $message) 

#source other scripts
. .\DovetailCommonFunctions.ps1

#validate input parameters
if ($caseId -eq "" -or $channel -eq "" -or $channel -eq "") 
{
   write-host "Missing Required parameters."
   write-host "Usage:"
   write-host
   write-host "SolutionMessage.ps1 solutionId channel (without the leading pound sign) message"
   write-host
   exit
}

$slackUrl="";

if ($channel -eq "dovetail")
{
	$slackUrl="https://hooks.slack.com/services/T0xx8LC/B0xxxPJ/LyYzT4YVxxx6M2Hq"
}

if ($slackUrl -eq ""){
  write-host
	write-host "Unknown channel: $channel";
	write-host
	exit
}
   
$ClarifyApplication = create-clarify-application; 
$ClarifySession = create-clarify-session $ClarifyApplication; 

$solution = get-row-for-table-by-unique-field "fc_user_solution_view" "id_number" $solutionId;
$probdesc = get-row-for-table-by-unique-field "probdesc" "id_number" $solutionId;

$URLBase = "https://agent.mycompany.com/support/"  

$solutionURLBase = $URLBase + "solutions/"  
$solutionUrl = $solutionURLBase + $solution.id_number;

$pretext = "Solution " + $solution.id_number;
# + " from " + $case.first_name + " " + $case.last_name + " at " + $case.site_name;
$fallback = $pretext + "; " + $solution.title + " " + $solutionURL;

$slackAttachment = New-Object -TypeName PSObject
Add-Member -InputObject $slackAttachment -MemberType NoteProperty -Name fallback -Value $fallback
Add-Member -InputObject $slackAttachment -MemberType NoteProperty -Name title -Value $solution.title;
Add-Member -InputObject $slackAttachment -MemberType NoteProperty -Name title_link -Value $solutionUrl;
Add-Member -InputObject $slackAttachment -MemberType NoteProperty -Name color -Value "good";

$availability = "Private"
if ($probdesc.public_ind -eq 1){
  $availability = "Public"
}

$field1 = New-Object -TypeName PSObject
Add-Member -InputObject $field1 -MemberType NoteProperty -Name title -Value "Availability";
Add-Member -InputObject $field1 -MemberType NoteProperty -Name value -Value $availability;
Add-Member -InputObject $field1 -MemberType NoteProperty -Name short -Value $false;

$field2 = New-Object -TypeName PSObject
Add-Member -InputObject $field2 -MemberType NoteProperty -Name title -Value "Type";
Add-Member -InputObject $field2 -MemberType NoteProperty -Name value -Value $probdesc.x_solution_type;
Add-Member -InputObject $field2 -MemberType NoteProperty -Name short -Value $false;

$fields = @($field1,$field2);

Add-Member -InputObject $slackAttachment -MemberType NoteProperty -Name fields -Value $fields;

$slackObject = [PSCustomObject]@{
		channel = "#$channel";
        text = "$message";
        username = "dovetail-bot";
}

$attachments = @($slackAttachment);

Add-Member -InputObject $slackObject -MemberType NoteProperty -Name attachments -Value $attachments;

$json = $slackObject | convertTo-json -Depth 4;
#$json;

Invoke-WebRequest -Uri $slackUrl -Method POST -Body $json > $null;

