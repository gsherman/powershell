param ([Parameter(Mandatory)]$inputFile, $dryRun=$true)

<#
Calling syntax:
	.\GetSolutionAttachmentDetails.ps1 input.xlsx 
#>

# Global Variables
$username = "dovetail-api";
$password="secret";

# url for get-solution API
# local
# $getSolutionUrl="https://default.lclhst.io/api/v1/solutions/";
# cloud
$getSolutionUrl="https://mytenant.dovetailnow.com/api/v1/solutions/";

$global:solutionAttachments = @();

<# Log Levels:
    OFF = 0
    ERROR = 1
    WARN = 2
    INFO = 3
    DEBUG = 4
#>
$global:logLevel=3;

# Giddyup

# build the log file name
$thisScript = (Get-Item $PSCommandPath ).Basename;
$FormattedDate = Get-Date -Format "yyyy-MM-dd-HH-mm-ss";
$logFile = "logs\$($thisScript)_$($FormattedDate).log";
$outputFile = "SolutionAttachments_$($FormattedDate).out";

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

function GetAttachmentData{
 param( $row, [int]$rowIndex )

	write-log " "  -Level "Info";
	write-log "--------------------------"  -Level "Info";
	write-log "Processing row $rowIndex" -Level "Debug"; 

	$credPair = "$($username):$($password)" 
	$encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($credPair)) 
	$requestHeaders = @{ 
		Authorization = "Basic $encodedCredentials" 
		Accept = "application/json"
	} 

	$solutionId = $row."Solution Id".toString().Trim();
	write-log "Getting data for solution $solutionId" -Level "Info"; 	

	$url = $getSolutionUrl + $solutionId;

	$response = Invoke-WebRequest -Uri $url  -Headers $requestHeaders
	$statusCode = $response.statuscode;
	write-log -Message "HTTP Request Success. HTTP Status Code: $statusCode"  -Level "Debug";
	$responseObject = $response | convertFrom-Json;

	$attachmentsArray= $responseObject.attachments;	
	$numAttachments = $attachmentsArray.length;
	
	write-log "Found $numAttachments attachments for solution $solutionId" -Level "Info"; 

	$owner = $row."Owner First Name".toString().Trim() + " " + $row."Owner Last Name".toString().Trim();	

	foreach ($attachment in $attachmentsArray){				
		$a = New-Object -TypeName PSObject -Property @{
			Id = $solutionId
			Title = $responseObject.title
			SolutionType = $row."Solution Type".toString().Trim()
			Owner= $owner
			OwnerLoginName= $responseObject.owner; 						
			File=$attachment.name
			ScanDate=$attachment.scanDate
			AgentUrl = $responseObject.agentUrl;
		}
		$global:solutionAttachments += $a;
	}

} #end function


##################
# Main
##################
If ($null -eq (Get-Module -ListAvailable -Name "ImportExcel")){
	Install-module ImportExcel;
}

write-log "Reading data from excel file: $($inputFile)" -Level "Info";
$excel = Import-Excel -Path $inputFile;

write-log "Successfully read data from excel file. Number of data rows: $($excel.count)" -Level "Info";
write-log "Row 1 is a header row" -Level "Info";

$rowCounter = 1;

foreach ($excelRow in $excel){
	$rowCounter = $rowCounter + 1;
	GetAttachmentData $excelRow $rowCounter;
}

$global:solutionAttachments | Export-Csv -Path attachments.csv