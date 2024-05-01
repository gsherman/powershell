<#
Calling syntax:

	.\create_and_updateusers.ps1 	

Outputs: 

 	* log file, in the \logs directory
#>

# WARNING! This is hacky. Super hacky. Best advice: don't use this script.

# Giddyup

<# Log Levels:
    OFF = 0
    ERROR = 1
    WARN = 2
    INFO = 3
    DEBUG = 4
#>
$global:logLevel=4;
$global:DovetailAgentCookie="COOKIE-VALUE-GOES-HERE";
$global:AgentUrl="https://qa3.dovetailtest.com/agent"

# initialize the pass/fail counters
$global:pass=0;
$global:fail=0;

$global:sites = New-Object System.Collections.Generic.List[System.Object]
$global:userId = "";


# build the log file name
$thisScript = (Get-Item $PSCommandPath ).Basename;
$FormattedDate = Get-Date -Format "yyyy-MM-dd-HH-mm-ss";
$logFile = "logs\$($thisScript)_$($FormattedDate).log";

# dateTime format
$dateTimeFormat="yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff'Z'";

# idMappings - map old case IDs to newly created case ID and save to a file for later use
$idMappings = @{}
$altID = '';
$idMappingFileName = "logs\idMappings_$($FormattedDate).json";

##########################
# Log to Console and to a File
##########################
function Write-Log
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Error","Warn","Info","Debug")]
        [string]$Level="Info"
    )

	if (!(Test-Path $logFile)) {
    	$NewLogFile = New-Item $logFile -Force -ItemType File
      }

    # Format Date for our Log File
    $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Write message to error, warning, or verbose pipeline and specify $LevelText
    switch ($Level) {
        'Error' {
	        	if ($global:logLevel -gt 0){
		            Write-Host 'ERROR:' $Message -ForegroundColor Red;
		            $LevelText = 'ERROR:'
		            "$FormattedDate $LevelText $Message" | Out-File -FilePath $logFile -Append
	            }
            }
        'Warn' {
	        	if ($global:logLevel -gt 1){
		            Write-Warning $Message
		            $LevelText = 'WARNING:'
		            "$FormattedDate $LevelText $Message" | Out-File -FilePath $logFile -Append
	            }            
            }
        'Info' {
				if ($global:logLevel -gt 2){
		            Write-Host $Message
		            $LevelText = 'INFO:'
		            "$FormattedDate $LevelText $Message" | Out-File -FilePath $logFile -Append
	            }
            }
         'Debug'{
				if ($global:logLevel -gt 3){
					Write-Host "DEBUG: $Message" -ForegroundColor Blue; 
					$LevelText = 'DEBUG:'
					"$FormattedDate $LevelText $Message" | Out-File -FilePath $logFile -Append
				}
         	}  

        } #end switch
        
} #end function


##################
# Process Response
##################
function ProcessResponse{
	param( $response)

	$status = $response.statuscode;
	$userId= $responseObject.entity.id;
	$global:userId = $userId;
	# write-log $userId;


	if ($status -eq 200){
		write-log -Message "Success.";
		$global:pass++;
	 }else{
	 	write-log -Message "Create User Failed" -Level "Error"
	 	$global:fail++;
	 }

}

function makeWebPostRequest{
	param ([string]$url, $jsonRequestBody )

	#write-log "makeWebPostRequest to $url"

$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
$session.Cookies.Add((New-Object System.Net.Cookie("dovetail.timezone.nag", "stop-nagging-me", "/", "qa3.dovetailtest.com")))
$session.Cookies.Add((New-Object System.Net.Cookie("__RequestVerificationToken", "REQUEST-VERIFICATION-TOKEN-GOES-HERE", "/", "qa3.dovetailtest.com")))
$session.Cookies.Add((New-Object System.Net.Cookie("DovetailCRMRememberedCulture", "en-US", "/", "qa3.dovetailtest.com")))

$session.Cookies.Add((New-Object System.Net.Cookie(".DOVETAIL_AGENT", $global:DovetailAgentCookie, "/", "qa3.dovetailtest.com")))

$response = Invoke-WebRequest  -UseBasicParsing -Uri "$url" `
-Method "POST" `
-WebSession $session `
-Headers @{
"authority"="qa3.dovetailtest.com"
  "method"="POST"
  "path"="/agent/api/users"
  "scheme"="https"
  "__requestverificationtoken"="REQUEST-VERIFICATION-TOKEN-GOES-HERE"
  "accept"="*/*"
  "accept-encoding"="gzip, deflate, br, zstd"
  "accept-language"="en-US,en;q=0.9"
  "cache-control"="no-cache"
  "origin"=$global:AgentUrl
  "pragma"="no-cache"
  "priority"="u=1, i"
  "referer"="$global:AgentUrl/agent/users/create"
  "sec-ch-ua"="`"Chromium`";v=`"124`", `"Brave`";v=`"124`", `"Not-A.Brand`";v=`"99`""
  "sec-ch-ua-mobile"="?0"
  "sec-ch-ua-platform"="`"Windows`""
  "sec-fetch-dest"="empty"
  "sec-fetch-mode"="cors"
  "sec-fetch-site"="same-origin"
  "sec-gpc"="1"
  "x-date"="2024-04-29T14:00:08-05:00"
  "x-requested-with"="XMLHttpRequest"
} `
-ContentType "application/json" `
-Body $jsonRequestBody

$response;

}
##################
# BuildRequestBody
##################
function BuildRequestBody{
 param( $randomUser)

	$JobTitleList='Payroll specialist','Talent acquisition specialist','HR analyst','Benefits administrator','Compensation specialist','HR Business Partner','HR Generalist','Retirement Specialist';
	$DepartmentList='General HR','Benefits','Accounting','Compensation','HRIS','Payroll','Talent Acquisition','Talent Management','Training'
	$WorkgroupList='HR','Benefits','People Services','Compensation','Recruiting','Analytics'
	$workCalendarIds='8f767904-76fd-49b4-aab1-b08f012cbe48','d243820b-a0c7-42e7-861d-b08f012cbe93','c144f9ca-6f4b-4fba-b0ed-b08f012cbeb8','81587a13-e7f5-4562-86d1-b08f012cbe5b'

	$queueIds = 'f452c254-58c7-4fbc-a162-b08f012ccd01','a85ca2a8-9f5b-4082-92d6-b10201177433','bd666668-0b09-4518-8997-b08f012ccd1d','b8db8341-594e-48ac-85e7-b08f012cccf3','2986c521-e9c3-4f41-8d61-b08f012ccd2b','f565bae2-b4b3-46a6-9996-b08f012ccd42','a245debd-587d-48a5-b5ec-b08f012ccd5a','8f933c57-176d-4d6c-9e26-b08f012ccd71','c71f5d32-5950-4c81-8ca7-b08f012ccd89'

	$user = @{
		status="Active";
	};
	$user.firstName = $randomUser.name.first;
	$user.lastName = $randomUser.name.last;
	$user.email = $randomUser.email;
	$user.username = $randomUser.login.username;

	$user.department= Get-Random -InputObject $DepartmentList;
	$user.jobTitle= Get-Random -InputObject $JobTitleList;
	$user.workgroup = Get-Random -InputObject $WorkgroupList;
	$user.defaultEmailTemplateId="";
	$user.defaultNoteTemplateId="";
	#$user.timezone=null;
	$user.calendarId= Get-Random -InputObject $workCalendarIds;
	$user.workCalendar = $user.calendarId;	
	$user.siteId=Get-Random -InputObject  $global:sites;

	# convert to json
	$jsonRequestBody = $user | convertto-json -Depth 5;
	$jsonRequestBody
}





function BuildUpdateUserQueuesRequestBody{
	param( $userId)
   
	   $queueIds = 'f452c254-58c7-4fbc-a162-b08f012ccd01','a85ca2a8-9f5b-4082-92d6-b10201177433','bd666668-0b09-4518-8997-b08f012ccd1d','b8db8341-594e-48ac-85e7-b08f012cccf3','2986c521-e9c3-4f41-8d61-b08f012ccd2b','f565bae2-b4b3-46a6-9996-b08f012ccd42','a245debd-587d-48a5-b5ec-b08f012ccd5a','8f933c57-176d-4d6c-9e26-b08f012ccd71','c71f5d32-5950-4c81-8ca7-b08f012ccd89'
   
	   $body = @{};
	   $body.id = $userId;
	   $body.isSupervisor = $false;

	   $queues = @();
	   $randomQueues = Get-Random -InputObject $queueIds -Count 3;
	   $queues+=$randomQueues;
	   $body.queueIds = $queues;
   
	   # convert to json
	   $jsonRequestBody = $body | convertto-json -Depth 5;
	   $jsonRequestBody
   }

   function UpdateUserQueues{

	$body = BuildUpdateUserQueuesRequestBody $global:userId;

	write-log " "  -Level "Info";
	write-log "--------------------------"  -Level "Info";

	write-log -Message $body.ToString() -Level "Debug"; 

	 try{

		$url = "$global:AgentUrl/api/users/$global:userId/queues/add";
		$response = makeWebPostRequest $url $body;

		$statusCode = $response.statuscode;

		write-log -Message "HTTP Request Success. HTTP Status Code: $statusCode"  -Level "Info";
		write-log $response -Level "Debug";
		$responseObject = $response | convertFrom-Json;
		# ProcessResponse $response;

	    }catch {		
		 	$response = $_.ErrorDetails.Message
		   	$err=$_.Exception;
			$statusCode = $err.Response.StatusCode.value__
			
			write-log -Message "HTTP UpdateQueues Request Failed. HTTP Status Code: $statusCode"  -Level "Error";
			if ($response) {write-log $response -Level "Debug";}

			if (($statusCode -eq 400) -and ($response)){
				ProcessResponse $response;
			}else{
				$global:fail++;
			}

	 } # end catch block

   }

   function BuildUpdateUserRolesRequestBody{
	param( $userId)
   
	   $roleIds = 'd156d3c2-75d9-45ef-bfa8-b1450142e51e'
	   # d156d3c2-75d9-45ef-bfa8-b1450142e51e = HR Agent
   
	   $body = @{};
	   $body.id = $userId;

	   $roles = @();
	   $randomRoles = Get-Random -InputObject $roleIds -Count 1;
	   $roles+=$randomRoles;
	   $body.roleIds = $roles;
   
	   # convert to json
	   $jsonRequestBody = $body | convertto-json -Depth 5;
	   $jsonRequestBody
   }
   
   function UpdateUserRoles{

	$body = BuildUpdateUserRolesRequestBody $global:userId;

	write-log " "  -Level "Info";
	write-log "--------------------------"  -Level "Info";

	write-log -Message $body.ToString() -Level "Debug"; 

	 try{

		$url = "$global:AgentUrl/api/users/$global:userId/roles/add";
		$response = makeWebPostRequest $url $body;

		$statusCode = $response.statuscode;

		write-log -Message "HTTP Request Success. HTTP Status Code: $statusCode"  -Level "Info";
		write-log $response -Level "Debug";
		$responseObject = $response | convertFrom-Json;
		# ProcessResponse $response;

	    }catch {		
		 	$response = $_.ErrorDetails.Message
		   	$err=$_.Exception;
			$statusCode = $err.Response.StatusCode.value__
			
			write-log -Message "HTTP UpdateRoles Request Failed. HTTP Status Code: $statusCode"  -Level "Error";
			if ($response) {write-log $response -Level "Debug";}

			if (($statusCode -eq 400) -and ($response)){
				ProcessResponse $response;
			}else{
				$global:fail++;
			}

	 } # end catch block

   }  

##################
# Create User
##################
function CreateUser{
 param( $randomUser )

	write-log " "  -Level "Info";
	write-log "--------------------------"  -Level "Info";

	$jsonRequestBody = BuildRequestBody $randomUser;
	write-log -Message $jsonRequestBody.ToString() -Level "Debug"; 

	try{

		$response = makeWebPostRequest "$global:AgentUrl/api/users" $jsonRequestBody;

		$statusCode = $response.statuscode;

		write-log -Message "HTTP Request Success. HTTP Status Code: $statusCode"  -Level "Info";
		write-log $response -Level "Debug";
		$responseObject = $response | convertFrom-Json;
		ProcessResponse $response;

	   }catch {		
		 	$response = $_.ErrorDetails.Message
		   	$err=$_.Exception;
			$statusCode = $err.Response.StatusCode.value__
			
			write-log -Message "HTTP Request Failed. HTTP Status Code: $statusCode"  -Level "Error";
			if ($response) {write-log $response -Level "Debug";}

			if (($statusCode -eq 400) -and ($response)){
				ProcessResponse $response;
			}else{
				$global:fail++;
			}

	 } # end catch block

} #end function


##################
# Main
##################

If ((Get-Module -ListAvailable -Name "ImportExcel") -eq $null){
	Install-module ImportExcel;
}

$excel = Import-Excel -Path "sites.xlsx";
foreach ($excelRow in $excel){
	$global:sites.Add( [string]$excelRow.SiteLinkID);
}


$numUsers = 10;
$counter = $numUsers;

while ($counter -gt 0){
	$randomUsers = Invoke-WebRequest -Uri "https://randomuser.me/api/" -Method Get 
	$randomUser = $randomUsers | convertfrom-json -Depth 5;

	CreateUser $randomUser.results[0] $sites; 
	UpdateUserQueues $global:userId;
	UpdateUserRoles $global:userId;

	$counter--;
}

 write-log "------------------------" -Level "Info";
 write-log "Pass: $($global:pass) of $($numUsers)" -Level "Info";
 write-log "Fail: $($global:fail) of $($numUsers)" -Level "Info";
