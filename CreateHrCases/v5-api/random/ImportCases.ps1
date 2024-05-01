<#
Calling syntax:

	.\importcases.ps1  	

Outputs: 

 	* log file, in the \logs directory
#>

# NOTES
# - this script is VERY dependent on the data that is in the tenant 
# - you'll have to set the values you have in your database for $PriorityList, $SeverityList, $CaseTypeList, etc.
# - ideally, the script would read those from the database using the List APIs. A task for another day
# - you'll also have  to edit the SetCustomFields function, to match the FAs you have in your tenant. 
# - to start a run, be sure to set $numCases, $condition (Open or Closed), $includeFAs, $dryRun
# - it's also dependent on having excel files of data 
#     -solutions (with a title column)
#     -users (with a username columnn)
#     -employees (with an employeeId column)
# - you can run reports and save to excel to create those excel files. be sure to edit the excel file to remove any extra rows at the top.


# TODO 
# - close date can't be in the future
# - read the list values from the database using the List APIs. A 

# Giddyup

# Settings
$numCases = 100;
$condition = 'Open' 
$includeFAs = $true;
$dryRun=$false;

$username = "dovetail-api";
$password="enter-secret-password-here";
$url="https://qa3.dovetailtest.com/api/v5/cases";

$failOnWarnings=$false;


<# Log Levels:
    OFF = 0
    ERROR = 1
    WARN = 2
    INFO = 3
    DEBUG = 4
#>
$global:logLevel=4;

# initialize the pass/fail counters
$global:pass=0;
$global:fail=0;
$global:employees = New-Object System.Collections.Generic.List[System.Object]
$global:titles = New-Object System.Collections.Generic.List[System.Object]
$global:users = New-Object System.Collections.Generic.List[System.Object]
$global:caseType=$null;

# build the log file name
$thisScript = (Get-Item $PSCommandPath ).Basename;
$FormattedDate = Get-Date -Format "yyyy-MM-dd-HH-mm-ss";
$logFile = "logs\$($thisScript)_$($FormattedDate).log";

# dateTime format
$dateTimeFormat="yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff'Z'";

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
	$warningsArray= $responseObject.warnings;
	$errorsArray= $responseObject.errors;
	$caseId = $null;

	if ($status -eq 201){
		$href= $responseObject.href;
		$id= $responseObject.id;
		$caseId= $responseObject.identifier;

		if ($dryRun -eq $true){
			write-log -Message "Success 201. Dry Run. " -Level "Info";
		}else{
			write-log -Message "Success 201. Newly created case id: $($caseId) " -Level "Info";
		}
		$global:pass++;		

	 }else{
	 	write-log -Message "Create Case Failed on row $rowIndex with status $status" -Level "Error"
	 	$global:fail++;
	 }

	foreach ($warning in $warningsArray){
		$message = "on row " + $rowIndex + " for field: " + $warning.field + " - " + $warning.description;
		write-log  -Message $message -Level "Warn";
	}

	foreach ($e in $errorsArray){
		$message = "on row " + $rowIndex + " for field: " + $e.field + " - " + $e.description;
		write-log -Message $message -Level "Error";
	}

	return $caseId;
}

##################
# BuildRequestBody
##################
function BuildRequestBody{
	param( [string]$condition, [bool]$includeFAs)

	# $apiKey = "1320ede31cec496c9bf960fd642ded14"
	# $randomText = Invoke-WebRequest -Uri "https://randommer.io/api/Text/LoremIpsum?loremType=business&type=paragraphs&number=3" -Method Get -Headers @{"X-Api-Key"=$apiKey}
	# $noteContent = $randomText.Content;
	# $caseNotes= $noteContent.Trim('"'); 

	$ipsums = 'song-lyrics','star-trek','corporate','movie-quotes','simpsons','seaside';
	$whichIpsum = Get-Random -InputObject  $ipsums;
	$url = "https://power-plugins.com/api/flipsum/ipsum/" + $whichIpsum + "?paragraphs=3&start_with_fixed=0"
	$randomText = Invoke-WebRequest -Uri $url
	$noteContent = $randomText | convertfrom-json
	$caseNotes = $noteContent[0] + $noteContent[1]+  $noteContent[2];

	$PriorityList='Urgent','High','Medium','Low';
	$SeverityList='Urgent','High','Medium','Low';
	$CaseTypeList = 'General HR','Benefits','Payroll','Policies','Training','Health','Talent Acquisition','Talent Management','Time and Attendance','Termination','HRIS','Employee Relations','Compensation';
	$OriginList = 'Manual Entry','Chat','Email In','Portal','API';
	$OpenStatusList = 'Open','On Hold','Waiting','Researching','In Progress';
	$QueueList = '','Benefits','General HR','Payroll','Compensation','HRIS';
	$LabelsList = 'Red','Brown','Orange','Lime','Green','Teal','Blue','Purple','Pink','Black','Crimson','Fuschia','Gold','Magenta','Indigo','Aqua','Navy','Tan','Silver','Slate'
	$TrueFalseList = 'TRUE','FALSE';

	if ($includeFAs){
		# case types with FAs
		$CaseTypeList = 'Separation','Payroll:Bonus','Compensation';
	}

	$case = @{  
		createEvents = "FALSE";
		employeeId = '';
	}

	$case.title=(Get-Random -InputObject  $global:titles).trim()
	$case.employeeId=Get-Random -InputObject  $global:employees;	
	$case.notes = $caseNotes;

	$case.severity = Get-Random -InputObject  $SeverityList;
	$case.priority = Get-Random -InputObject  $PriorityList;
	$case.caseType = Get-Random -InputObject  $CaseTypeList;
	$global:caseType = $case.caseType;

	$case.origin   = Get-Random -InputObject  $OriginList;
	$case.sensitive = Get-Random -InputObject $TrueFalseList;
	$case.portalCaseType = "GeneralHR";
	$case.availableInPortal = Get-Random -InputObject $TrueFalseList;	

	$labels = Get-Random -InputObject $LabelsList -Count (Get-Random -Minimum 1 -Maximum 5)
	$case.labels = @($labels)

	$case.originatorUserName = Get-Random -InputObject  $global:users;
	$case.ownerUserName = Get-Random -InputObject  $global:users;

	# random concerning employee, for some cases
	if (Get-Random -InputObject ([bool]$True,[bool]$False)) {
		$case.concerningEmployeeId = Get-Random -InputObject  $global:employees;
	 }


	if ($condition -eq 'Closed'){
		$case.condition = "Closed";
		$case.status = "Closed";
		$case.closeNotes = "Done, and done."
		$case.closeResolution = "Closed - Resolved"
		# $case.queue= NULL, AS THE CASE IS CLOSED

		# random create date in the past 4 years
		$Start = Get-Date '01.01.2020'
		$End = Get-Date '04.30.2024'
		$Random = Get-Random -Minimum $Start.Ticks -Maximum $End.Ticks
		$dt= [datetime]$Random
		$case.createDate = $dt.ToUniversalTime().ToString($dateTimeFormat); 

		# TODO - close date can't be in the future, which can happen here
		# maybe just do a random close date between the createDate and now?

		# random close date, after the create date 
		$numDaysToClose = Get-Random -Minimum 1 -Maximum 30;
		$closeDate = $dt.AddDays($numDaysToClose);
		$randomMinutes = Get-Random -Minimum 10 -Maximum 500;		
		$closeDate = $closeDate.AddMinutes($randomMinutes);		
		$closeDate = $closeDate.AddSeconds($randomMinutes);	
		$case.closeDate = $closeDate.ToUniversalTime().ToString($dateTimeFormat); 
	} else{
		$case.condition = "Open";
		$case.status = Get-Random -InputObject $OpenStatusList;
		$case.queue= Get-Random -InputObject $QueueList;

		# random create date in the near past (this year)
		$Start = Get-Date '01.01.2024'
		$End = Get-Date '05.01.2024'
		$Random = Get-Random -Minimum $Start.Ticks -Maximum $End.Ticks
		$dt= [datetime]$Random
		$case.createDate = $dt.ToUniversalTime().ToString($dateTimeFormat); 		
	}


	$caseJson = $case | convertto-json -Depth 5;

	# set the options
	$case.dryRun = $dryRun;
	$case.failOnWarnings = $failOnWarnings;

	# convert to json
	$jsonRequestBody = $case | convertto-json -Depth 5;
	$jsonRequestBody
}

##################
# Set Custrom Fields (FAs)
##################
function SetCustomFields{
	param( [string]$caseId)

	write-log -Message "setting custom fields..."  -Level "Info";

	$FAurl = $url +  "/" + $caseId + "/flexibleattributes"
	$FAbody = @{}

	$Start = Get-Date '01.01.2024'
	$End = Get-Date '05.01.2024'
	$Random = Get-Random -Minimum $Start.Ticks -Maximum $End.Ticks
	$dt= [datetime]$Random

	$FA1 = @{  
		key = "date";
		value = $dt.ToUniversalTime().ToString($dateTimeFormat); 
	}

	if ($global:caseType -eq "Separation"){
		$ReasonList = 'Voluntary','Involuntary','Dismissal','Retirement','Layoffs';
		$FA2 = @{  
			key = "reason";
			value = Get-Random -InputObject  $ReasonList; 
			}
		$FAbody.FlexibleAttributes = @($FA1,$FA2)		
	}

	if ($global:caseType -eq "Payroll:Bonus"){
		$randomDecimal = get-random -Maximum ([Decimal]50000);
		$amount = [math]::Round($randomDecimal,2);

		$FA2 = @{  
			key = "amount";
			value = $amount;
			}
		$FAbody.FlexibleAttributes = @($FA1,$FA2)
	}

	if ($global:caseType -eq "Compensation"){
		$CompensationChangeList = 'Increase','Decrease';

		$randomDecimal = get-random -Maximum ([Decimal]150000);
		$amount = [math]::Round($randomDecimal,2);

		$FA2 = @{  
			key = "new_salary";
			value = $amount;
			}
		$FA3 = @{  
			key = "type";
			value = Get-Random -InputObject  $CompensationChangeList; 
			}			
		$FAbody.FlexibleAttributes = @($FA1,$FA2,$FA3)
	}
	
	$FAbodyJson = $FAbody | convertto-json -Depth 5;

	write-log -Message "PUT FlexibleAttributes to $FAurl "  -Level "Info";
	write-log -Message $FAbodyJson.ToString() -Level "Debug"; 

	if ($dryRun){
		exit;
	}

	try{
		$response = Invoke-WebRequest -Uri $FAurl -Method Put -Body $FAbodyJson -ContentType "application/json" -Headers $requestHeaders 
		$statusCode = $response.statuscode;

		write-log -Message "PUT FlexibleAttributes HTTP Request Success. HTTP Status Code: $statusCode"  -Level "Info";
		write-log $response -Level "Debug";

		$responseObject = $response | convertFrom-Json;
		$caseId = ProcessResponse $response;

	   }catch {
		 	$response = $_.ErrorDetails.Message
		   	$err=$_.Exception;
			$statusCode = $err.Response.StatusCode.value__
			
			write-log -Message "PUT FlexibleAttributes HTTP Request Failed. HTTP Status Code: $statusCode"  -Level "Error";
			if ($response) {write-log $response -Level "Debug";}
	 } # end catch block

}

##################
# Create Case
##################
function CreateCase{
 param( [string]$condition, [bool]$includeFAs)

	write-log " "  -Level "Info";
	write-log "--------------------------"  -Level "Info";
	write-log "Processing row $rowIndex" -Level "Info"; 

	$jsonRequestBody = BuildRequestBody $condition $includeFAs;
	write-log -Message $jsonRequestBody.ToString() -Level "Debug"; 

	$credPair = "$($username):$($password)" 
	$encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($credPair)) 
	$requestHeaders = @{ 
		Authorization = "Basic $encodedCredentials" 
		Accept = "application/json"
	} 

	try{
		$response = Invoke-WebRequest -Uri $url -Method Post -Body $jsonRequestBody -ContentType "application/json" -Headers $requestHeaders 
		$statusCode = $response.statuscode;

		write-log -Message "HTTP Request Success. HTTP Status Code: $statusCode"  -Level "Info";
		write-log $response -Level "Debug";

		$responseObject = $response | convertFrom-Json;
		$caseId = ProcessResponse $response;

	   }catch {

		 	$response = $_.ErrorDetails.Message
		   	$err=$_.Exception;
			$statusCode = $err.Response.StatusCode.value__
			
			write-log -Message "HTTP Request Failed. HTTP Status Code: $statusCode"  -Level "Error";
			if ($response) {write-log $response -Level "Debug";}

			if (($statusCode -eq 400) -and ($response)){
				$responseObject = $response | convertFrom-Json;
				$caseId = ProcessResponse $response;
			}else{
				$global:fail++;
			}

	 } # end catch block

	 if ($includeFAs){
		if ($caseId -ne $null){
			SetCustomFields $caseId;
		}
	 }
} #end function



##################
# Main
##################
If ((Get-Module -ListAvailable -Name "ImportExcel") -eq $null){
	Install-module ImportExcel;
}

# load up list of employees IDs from excel
$excel = Import-Excel -Path "employees.xlsx";
foreach ($excelRow in $excel){
	$global:employees.Add( [string]$excelRow.EmployeeId);
}

# load up list of users from excel
$excel = Import-Excel -Path "users.xlsx";
foreach ($excelRow in $excel){
	$global:users.Add( [string]$excelRow.Username);
}

# load up solution titles from excel
$excel = Import-Excel -Path "solutions.xlsx";
foreach ($excelRow in $excel){
	$global:titles.Add( [string]$excelRow.Title);
}

$counter = $numCases;
while ($counter -gt 0){
	CreateCase $condition $includeFAs;
	$counter--;
}

 write-log "------------------------" -Level "Info";
 write-log "Pass: $($global:pass) of $($numCases)" -Level "Info";
 write-log "Fail: $($global:fail) of $($numCases)" -Level "Info";
