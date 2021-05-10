param (
	[Parameter(Mandatory)]$inputFile, 
	[int][ValidateRange(2, [int]::MaxValue)]$onlyProcessRowNumber
)

<#
Powershell script to create new Flexible Attributes
It reads data from an Excel file, and creates FAs for a given case type
It will either process all the rows, or only 1 row (if the onlyProcessRowNumber input param is specified)

Excel sheet format:
Header row, with following headings:
	* caseTypeKey (required)
	* label (required)
	* key  (required)	
	* dataType (required, one of: string, bool, datetime, decimal, int, list)	
	* listName (optional)
	* defaultValue (optional)	
	* required (optional, True or False. defaults to False)		
Then, one row per FA to be created.

Calling syntax:
	.\createFAs.ps1 input.xlsx 							Process all rows from Excel
	.\createFAs.ps1 input.xlsx -onlyProcessRowNumber 2	Only process one row from Excel					

Outputs: 
 	* log file, in the \logs directory
#>


# Settings
$username = "administrator"; # Must be a RootAdmin in order to create FAs
$password="administrator";
$url="http://localhost/api/v1/flexible-attributes/cases";

<# Log Levels:
    OFF = 0
    ERROR = 1
    WARN = 2
    INFO = 3
    DEBUG = 4
#>
$global:logLevel=4;


# Giddyup

# initialize the pass/fail counters
$global:pass=0;
$global:fail=0;

# build the log file name
$thisScript = (Get-Item $PSCommandPath ).Basename;
$FormattedDate = Get-Date -Format "yyyy-MM-dd-HH-mm-ss";
$logFile = "logs\$($thisScript)_$($FormattedDate).log";

# dateTime format
$dateTimeFormat="yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff'Z'";

###############################
# Log to Console and to a File
###############################
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
	# $warningsArray= $responseObject.warnings;
	$errorsArray= $responseObject.errors;

	if ($status -eq 201){
		write-log -Message "Success 201." -Level "Info";
		$global:pass++;

	 }else{
	 	write-log -Message "Create Flexible Attribute Failed on row $rowIndex with status $status" -Level "Error"
	 	$global:fail++;
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

	$fa = @{  
		required = "FALSE";
#		listName = '';
#		defaultValue = '';
	}

	if ($row.caseTypeKey) { $fa.caseType = $row.caseTypeKey.toString(); }
	if ($row.label) {$fa.label = $row.label.toString(); }
	if ($row.key) { $fa.key = $row.key.toString(); }
	if ($row.dataType) { $fa.dataType = $row.dataType.toString(); }
	if ($row.listName) { $fa.listName = $row.listName.toString(); }	
	if ($row.defaultValue) { $fa.defaultValue = $row.defaultValue.toString(); }
	if ($row.required) { $fa.required = $row.required.toString(); }	

	# convert to json
	$jsonRequestBody = $fa | convertto-json;
	$jsonRequestBody
}


##################
# Create FA
##################
function CreateFA{
 param( $row, [int]$rowIndex )

	write-log " "  -Level "Info";
	write-log "--------------------------"  -Level "Info";
	write-log "Processing row $rowIndex" -Level "Info"; 

	$jsonRequestBody = BuildRequestBody $row $rowIndex;
	write-log -Message $jsonRequestBody.ToString() -Level "Debug"; 

	$credPair = "$($username):$($password)" 
	$encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($credPair)) 
	$requestHeaders = @{ 
		Authorization = "Basic $encodedCredentials" 
		Accept = "application/json"
	} 

	try{
		$response = Invoke-WebRequest -Uri $url -Method Post -Body $jsonRequestBody -ContentType "application/json; charset=utf-8" -Headers $requestHeaders 
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
				$responseObject = $response | convertFrom-Json;
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

write-log "Reading data from excel file: $($inputFile)" -Level "Info";

$excel = Import-Excel -Path $inputFile;

write-log "Successfully read data from excel file. Number of data rows: $($excel.count)" -Level "Info";
write-log "Row 1 is a header row" -Level "Info";

$numberOfRowsProcessed = 0;

if ($onlyProcessRowNumber){
	write-log "Only processing row $onlyProcessRowNumber" -Level "Info";
	$numberOfRowsProcessed = $numberOfRowsProcessed + 1;
	CreateFA $excel[$onlyProcessRowNumber -2] $onlyProcessRowNumber; # Minus 2, as it's a zero based array, starting after the header row
}else{
	$rowIndex = 1;
	foreach ($excelRow in $excel){
		$rowIndex = $rowIndex + 1;	
		$numberOfRowsProcessed = $numberOfRowsProcessed + 1;
		CreateFA $excelRow $rowIndex;
	}
}

 write-log "------------------------" -Level "Info";
 write-log "Pass: $($global:pass) of $($numberOfRowsProcessed)" -Level "Info";
 write-log "Fail: $($global:fail) of $($numberOfRowsProcessed)" -Level "Info";
