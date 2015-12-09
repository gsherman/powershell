##############################################################################
## SCRIPT:	CaseMessage.ps1
##
## PURPOSE:	This PowerShell script posts a case message to slack
##
## Call this script like: 
##		CaseMessage.ps1 caseId channel(without the leading pound sign) message
##
##############################################################################

param([string] $caseId, [string] $channel, [string] $message) 

$URLBase = "https://agent.mycompany.com/support/"  
$mobileURLBase = "http://tinyurl.com/abc123/"
$logoUrl = "https://mycompany.com/logo.jpg";

#source other scripts
. .\DovetailCommonFunctions.ps1

#validate input parameters
if ($caseId -eq "" -or $channel -eq "" -or $channel -eq "") 
{
   write-host "Missing Required parameters."
   write-host "Usage:"
   write-host
   write-host "CaseMessage.ps1 caseId channel (without the leading pound sign) message"
   write-host
   exit
}

$slackUrl="";

if ($channel -eq "dovetail")
{
	$slackUrl="https://hooks.slack.com/services/T0xxxR8LC/B0xxxxPJ/LyYzTxxxxxxb6M2Hq"
}

if ($slackUrl -eq ""){
  write-host
	write-host "Unknown channel: $channel";
	write-host
	exit
}
   
$ClarifyApplication = create-clarify-application; 
$ClarifySession = create-clarify-session $ClarifyApplication; 

$case = get-row-for-table-by-unique-field "fc_user_case_view" "id_number" $caseId;

$caseURLBase = $URLBase + "cases/"  
$caseUrl = $caseURLBase + $case.id_number;
$mobileCaseUrl = $mobileURLBase + $case.id_number;

$pretext = "Case " + $case.id_number + " from " + $case.first_name + " " + $case.last_name + " at " + $case.site_name;
$fallback = $pretext + "; " + $case.title + " " + $caseURL;

$slackAttachment = New-Object -TypeName PSObject
Add-Member -InputObject $slackAttachment -MemberType NoteProperty -Name fallback -Value $fallback
Add-Member -InputObject $slackAttachment -MemberType NoteProperty -Name title -Value $case.title;
Add-Member -InputObject $slackAttachment -MemberType NoteProperty -Name title_link -Value $caseUrl;
Add-Member -InputObject $slackAttachment -MemberType NoteProperty -Name color -Value "warning";
Add-Member -InputObject $slackAttachment -MemberType NoteProperty -Name pretext -Value $pretext;

#Case History
#$caseRecord = get-row-for-table-by-unique-field "case" "id_number" $caseId;
#Add-Member -InputObject $slackAttachment -MemberType NoteProperty -Name text -Value $caseRecord.case_history;

Add-Member -InputObject $slackAttachment -MemberType NoteProperty -Name thumb_url -Value $logoUrl;

$field1 = New-Object -TypeName PSObject
Add-Member -InputObject $field1 -MemberType NoteProperty -Name title -Value "Severity";
Add-Member -InputObject $field1 -MemberType NoteProperty -Name value -Value $case.severity;
Add-Member -InputObject $field1 -MemberType NoteProperty -Name short -Value $false;

$queue = $case.queue;
if ($queue.length -le 1){$queue="None";}

$field2 = New-Object -TypeName PSObject
Add-Member -InputObject $field2 -MemberType NoteProperty -Name title -Value "Queue";
Add-Member -InputObject $field2 -MemberType NoteProperty -Name value -Value $queue;
Add-Member -InputObject $field2 -MemberType NoteProperty -Name short -Value $false;

$field3 = New-Object -TypeName PSObject
Add-Member -InputObject $field3 -MemberType NoteProperty -Name title -Value "Mobile Link";
Add-Member -InputObject $field3 -MemberType NoteProperty -Name value -Value $mobileCaseUrl;
Add-Member -InputObject $field3 -MemberType NoteProperty -Name short -Value $false;

$fields = @($field1,$field2,$field3);

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

