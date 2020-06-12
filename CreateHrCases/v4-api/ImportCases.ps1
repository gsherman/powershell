param ([Parameter(Mandatory)]$inputFile, $dryRun=$true)

<#
Calling syntax:

	.\importcases.ps1 input.xlsx 							Implicit dryRun = true
	.\importcases.ps1 input.xlsx true						Explicit dryRun = true
	.\importcases.ps1 input.xlsx false 						Explicit dryRun = false
	.\ImportCases.ps1 -inputfile 1-cases.xlsx -dryrun true	Using named parameters

	Notes:
	* this assumes the create-date and close-date in the excel file are in the local time zone (where PS is being run), which may not be true. we may need to be more robust here
	* this script imports cases one at a time. it does not batch them. maybe add support later for creating a batch of cases. multiple cases in the Cases array. or maybe that's a different script?
#>


$failOnWarnings=$false;
$username = "dovetail-api";
$password="letmein";
$url="http://localhost/api/v5/cases";

<# Log Levels:
    OFF = 0
    ERROR = 1
    WARN = 2
    INFO = 3
    DEBUG = 4
#>
$global:logLevel=4;

# build the log file name
$thisScript = (Get-Item $PSCommandPath ).Basename;
$FormattedDate = Get-Date -Format "yyyy-MM-dd-HH-mm-ss";
$logFile = "$($thisScript)_$($FormattedDate).log";

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

	$resultsIndex=0;
	$status= $responseObject.results[$resultsIndex].status;
	$correlationId= $responseObject.results[$resultsIndex].correlationId;
	$warningsArray= $responseObject.results[$resultsIndex].warnings;
	$errorsArray= $responseObject.results[$resultsIndex].errors;

	if ($status -eq 201){
		$href= $responseObject.results[$resultsIndex].href;
		$id= $responseObject.results[$resultsIndex].id;
		$caseId= $responseObject.results[$resultsIndex].identifier;

		if ($dryRun -eq $true){
			write-log -Message "Success 201. Dry Run. For correlationId $($correlationId)" -Level "Info";
		}else{
			write-log -Message "Success 201. Newly created case id: $($caseId) for correlationId $($correlationId)" -Level "Info";
		}

	}else{
		write-log -Message "Create Case Failed on row $rowIndex for correlationId $correlationId with status $status" -Level "Error"
	}

	foreach ($warning in $warningsArray){
		$message = "on row " + $rowIndex + " for field: " + $warning.field + " - " + $warning.description;
		write-log  -Message $message -Level "Warn";
	}

	foreach ($e in $errorsArray){
		$message = "on row " + $rowIndex + " for field: " + $e.field + " - " + $e.description;
		write-log -Message $message -Level "Error";
	}

}

##################
# BuildRequestBody
##################
function BuildRequestBody{
 param( $row, [int]$rowIndex )

	$case = @{  
		createEvents = "FALSE";
		correlationID = "excel_row_" + $rowIndex;
		employeeId = '';
		notes = '';
	}

	if ($row.title) { $case.Title = $row.title.toString(); }
	if ($row.employeeId) { $case.employeeId = $row.employeeId.toString(); }
	if ($row.notes) { $case.notes = $row.notes.toString(); }
	if ($row.concerningEmployeeId) { $case.concerningEmployeeId = $row.concerningEmployeeId.toString(); }
	if ($row.ID) { $case.alternateId = $row.ID.toString(); }
	if ($row.severity) { $case.severity = $row.severity.toString(); }
	if ($row.priority) { $case.priority = $row.priority.toString(); }
	if ($row.condition) { $case.condition = $row.condition.toString(); }
	if ($row.status) { $case.status = $row.status.toString(); }
	if ($row.caseType) { $case.caseType = $row.caseType.toString(); }
	if ($row.portalCaseType) { $case.portalCaseType = $row.portalCaseType.toString(); }
	if ($row.origin) { $case.origin = $row.origin.toString(); }
	if ($row.queue) { $case.queue = $row.queue.toString(); }
	if ($row.originatorUserName) { $case.originatorUserName = $row.originatorUserName.toString(); }
	if ($row.ownerUserName) { $case.ownerUserName = $row.ownerUserName.toString(); }
	if ($row.availableInPortal) { $case.availableInPortal = $row.availableInPortal.toString(); }
	if ($row.sensitive) { $case.sensitive = $row.sensitive.toString(); }
	if ($row.createEvents) { $case.createEvents = $row.createEvents.toString(); }	
	if ($row.closeNotes) { $case.closeNotes = $row.closeNotes.toString(); }
	if ($row.closeResolution) { $case.closeResolution = $row.closeResolution.toString(); }

	# Convert CreateDate into UTC and into an ISO-8601 format
	# this assumes the date is in the file the local time zone (where PS is being run), which may not be true. we may need to be more robust here
	if ($row.createDate) { 
	 	$dt = Get-Date($row.createDate);
		$case.createDate = $dt.ToUniversalTime().ToString("s") + "Z"; 
	}

	# Convert CloseDate into UTC and into an ISO-8601 format
	# this assumes the date is in the file the local time zone (where PS is being run), which may not be true. we may need to be more robust here
	if ($row.closeDate) { 
	 	$dt = Get-Date($row.closeDate);
		$case.closeDate = $dt.ToUniversalTime().ToString("s")  + "Z"; 
	}

	# Turn labels into an array
	if ($row.labels) { 
		$labelsArray = @( $row.labels.toString().Split(",").Trim() );		
		$case.labels = $labelsArray;		
	}			

	# build up the body of the request 
	# set the global variables

	$body = @{
		dryRun = $dryRun;  
		failOnWarnings = $failOnWarnings;
		};

	$caseArray = @($case);	
	$body.Cases = $caseArray;

	# convert to json
	$jsonRequestBody = $body | convertto-json -Depth 5;
	$jsonRequestBody
}


##################
# Create Case
##################
function CreateCase{
 param( $row, [int]$rowIndex )

	write-log "--------------------------"  -Level "Info";
	write-log "Processing row $rowIndex" -Level "Info"; 

	# Validate Create Date
	if ($row.createDate) { 
		[datetime]$result = New-Object DateTime;

		if (!([DateTime]::TryParse($row.createDate,[ref]$result))){
			$createDate = $row.createDate;
			write-log -Message "Invalid createDate: $($createDate) on row $($rowIndex)" -Level "Error"; 
			return;
		}
	}

	# Validate Close Date
	if ($row.closeDate) { 
		[datetime]$result = New-Object DateTime;

		if (!([DateTime]::TryParse($row.closeDate,[ref]$result))){
			$closeDate = $row.closeDate;
			write-log -Message "Invalid closeDate: $($closeDate) on row $($rowIndex)" -Level "Error"; 
			return;
		}
	}

	$jsonRequestBody = BuildRequestBody $row $rowIndex;

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
		ProcessResponse $response;

	   }catch {

		 	$response = $_.ErrorDetails.Message
		   	$err=$_.Exception;
			$statusCode = $err.Response.StatusCode.value__

			write-log -Message "HTTP Request Failed. HTTP Status Code: $statusCode"  -Level "Error";
			if ($response) {write-log $response -Level "Debug";}

			# 407 - multi-status, some cases succeeded, some failed
			if (($statusCode -eq 407) -or ($statusCode -eq 400)){
				$responseObject = $response | convertFrom-Json;
				ProcessResponse $response;
			}else{
				write-log $response -Level "Debug";
			}

	 } # end catch block

} #end function



##################
# Main
##################
If ((Get-Module -ListAvailable -Name "ImportExcel") -eq $null){
	Install-module ImportExcel;
}

write-log "Reading data from excel file: $($inputFile)" -Level "Info";

$excel = Import-Excel -Path $inputFile;

write-log "Successfully read data from excel file. Number of data rows: $($excel.count)" -Level "Info";
write-log "Row 1 is a header row" -Level "Info";

$rowCounter = 1;

foreach ($excelRow in $excel){
	$rowCounter = $rowCounter + 1;
	CreateCase $excelRow $rowCounter;
}
