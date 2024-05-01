param ([Parameter(Mandatory)]$inputFile, $dryRun=$true)

<#
Calling syntax:

	.\importsites.ps1 input.xlsx 		

Outputs: 

 	* log file, in the \logs directory
#>

# WARNING! This is hacky. Super hacky. Best advice: don't use this script.


# Giddyup

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

# initialize the pass/fail counters
$global:pass=0;
$global:fail=0;

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
	$correlationId= $responseObject.correlationId;
	$warningsArray= $responseObject.warnings;
	$errorsArray= $responseObject.errors;

	if ($status -eq 201){
		$href= $responseObject.href;
		$id= $responseObject.id;
		$caseId= $responseObject.identifier;

		if ($dryRun -eq $true){
			write-log -Message "Success 201. Dry Run. For correlationId $($correlationId)" -Level "Info";
		}else{
			write-log -Message "Success 201. Newly created case id: $($caseId) for correlationId $($correlationId)" -Level "Info";
			$idMappings[$altID] = $caseId;
		}

		$global:pass++;

	 }else{
	 	write-log -Message "Create Case Failed on row $rowIndex for correlationId $correlationId with status $status" -Level "Error"
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

}

##################
# BuildRequestBody
##################
function BuildRequestBody{
 param( $row, [int]$rowIndex )

	$site = @{
		addressType="General";
		siteType= "Research and Development";
		status="Active";
		organizationId = "32cff24c-1483-48b2-a2a8-b1610132d39e";
	};

	$address=@{}

	if ($row.Address1) { $address.address1 = $row.Address1.toString(); }
	if ($row.State) { $address.stateOrProvince = $row.State.toString(); }
	if ($row.Country) { $address.country = $row.Country.toString(); }
	if ($row.PostalCode) { $address.postalCode = $row.PostalCode.toString(); }
	if ($row.Timezone) { $address.timeZoneId = $row.Timezone.toString(); }
	if ($row.State) { $address.stateOrProvince = $row.State.toString(); }
	if ($row.City) { $address.city = $row.City.toString(); }
	$address.addressType="General";
	$site.address=$address;

	$organization=@{}
	$organization.value= "32cff24c-1483-48b2-a2a8-b1610132d39e";
	$organization.label="1002 - Ceridian (Partner)";
	$site.organization=$organization;

	if ($row.Identifier) { $site.identifier = $row.Identifier.toString(); }
	if ($row.Name) { $site.name = $row.Name.toString(); }
	if ($row.Address1) { $site.address1 = $row.Address1.toString(); }
	if ($row.City) { $site.city = $row.City.toString(); }
	if ($row.State) { $site.stateOrProvince = $row.State.toString(); }
	if ($row.Country) { $site.country = $row.Country.toString(); }
	if ($row.PostalCode) { $site.postalCode = $row.PostalCode.toString(); }
	if ($row.Timezone) { $site.timeZoneId = $row.Timezone.toString(); }
	#if ($row.OrgName) { $site.organizationId = $row.OrgName.toString(); }
	if ($row.WorkCalendar) { $site.calendar = $row.WorkCalendar.toString(); }

	# convert to json
	$jsonRequestBody = $site | convertto-json -Depth 5;
	$jsonRequestBody
}


##################
# Create Site
##################
function CreateSite{
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


$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
$session.Cookies.Add((New-Object System.Net.Cookie("dovetail.timezone.nag", "stop-nagging-me", "/", "qa3.dovetailtest.com")))
$session.Cookies.Add((New-Object System.Net.Cookie("__RequestVerificationToken", "REQUEST-VERIFICATION-TOKEN-GOES-HERE", "/", "qa3.dovetailtest.com")))
$session.Cookies.Add((New-Object System.Net.Cookie("DovetailCRMRememberedCulture", "en-US", "/", "qa3.dovetailtest.com")))
$session.Cookies.Add((New-Object System.Net.Cookie(".DOVETAIL_AGENT", "COOKIE-VALUE-GOES-HERE", "/", "qa3.dovetailtest.com")))

$response = Invoke-WebRequest  -UseBasicParsing -Uri "https://qa3.dovetailtest.com/agent/api/sites" `
-Method "POST" `
-WebSession $session `
-Headers @{
"authority"="qa3.dovetailtest.com"
  "method"="POST"
  "path"="/agent/api/sites"
  "scheme"="https"
  "__requestverificationtoken"="REQUEST-VERIFICATION-TOKEN-GOES-HERE"
  "accept"="*/*"
  "accept-encoding"="gzip, deflate, br, zstd"
  "accept-language"="en-US,en;q=0.9"
  "cache-control"="no-cache"
  "origin"="https://qa3.dovetailtest.com"
  "pragma"="no-cache"
  "priority"="u=1, i"
  "referer"="https://qa3.dovetailtest.com/agent/sites/create"
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

$rowCounter = 1;

foreach ($excelRow in $excel){
	$rowCounter = $rowCounter + 1;
	CreateSite $excelRow $rowCounter;
}

 write-log "------------------------" -Level "Info";
 write-log "Pass: $($global:pass) of $($excel.count)" -Level "Info";
 write-log "Fail: $($global:fail) of $($excel.count)" -Level "Info";
