param ([Parameter(Mandatory)]$inputFile, [Parameter(Mandatory)]$fromQueue, [Parameter(Mandatory)]$toQueue)

<#

Context: https://support.dovetailsoftware.com/agent/support/cases/24846

Calling syntax:

	.\analysis.ps1 -inputfile cases.xlsx fromQueueName toQueueName
	.\Analysis.ps1 .\queued_cases_since_last_year.xlsx "General HR" "Payroll"

	Output Example: 

	Number of cases forwarded to the Payroll queue : 3 
	Number of cases forwarded to the Payroll queue from the General HR queue: 2 
	Number of cases dispatched to the Payroll queue : 5 
	Number of cases dispatched to the General HR queue : 31 
	Number of cases dispatched to both queues (General HR and Payroll) : 2 
	Number of cases dispatched to the General HR queue and then later dispatched to the Payroll queue) : 1 

#>


# Giddyup

<# Log Levels:
    OFF = 0
    ERROR = 1
    WARN = 2
    INFO = 3
    DEBUG = 4
#>
$global:logLevel=4;

# whether to write the data arrays to the output or not. 
$writeData = $false;

# initialize the arrays
$global:casesForwardedTo = @();
$global:casesForwardedFromTo = @();
$global:casesDispatchedToTheFromQueue = @();
$global:casesDispatchedToTheToQueue = @();
$global:casesDispatchedToBothQueues = @();
$global:casesDispatchedToTheFromQueueAndThenLaterToTheToQueue= @();

# build the log file name
$thisScript = (Get-Item $PSCommandPath ).Basename;
$FormattedDate = Get-Date -Format "yyyy-MM-dd-HH-mm-ss";
$logFile = "logs\$($thisScript)_$($FormattedDate).log";

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
# ProcessExcelRow
##################
function ProcessExcelRow{
 param( $row, [int]$rowIndex )

if ($row.Action -eq "Forward" -and $row.ToQueue -eq $toQueue) { $global:casesForwardedTo+= $row.CaseID; }
if ($row.Action -eq "Forward" -and $row.ToQueue -eq $toQueue -and $row.FromQueue -eq $fromQueue) { $global:casesForwardedFromTo+= $row.CaseID; }
if ($row.Action -eq "Dispatch" -and $row.ToQueue -eq $toQueue) { $global:casesDispatchedToTheToQueue+= $row.CaseID; }
if ($row.Action -eq "Dispatch" -and $row.ToQueue -eq $fromQueue) { $global:casesDispatchedToTheFromQueue+= $row.CaseID; }

# Assuming the data is in order of WorkflowLog.Created (When), ascending
if ($row.Action -eq "Dispatch" -and $row.ToQueue -eq $toQueue -and ($global:casesDispatchedToTheFromQueue -contains $row.CaseID) ) 
	{ $global:casesDispatchedToTheFromQueueAndThenLaterToTheToQueue+= $row.CaseID; }
}


##################
# Main
##################
If ($null -eq (Get-Module -ListAvailable -Name "ImportExcel")) {
	Install-module ImportExcel;
}

write-log "Reading data from excel file: $($inputFile)" -Level "Info";

$excel = Import-Excel -Path $inputFile;

write-log "Successfully read data from excel file. Number of data rows: $($excel.count)" -Level "Info";
write-log "Rows 1 is a header row" -Level "Info";

$rowCounter = 1;

foreach ($excelRow in $excel){
	$rowCounter = $rowCounter + 1;
	ProcessExcelRow $excelRow $rowCounter;
}

write-log "------------------------" -Level "Info";

$casesForwardedTo = $global:casesForwardedTo | Sort-Object | Get-Unique;
$casesForwardedFromTo = $global:casesForwardedFromTo | Sort-Object | Get-Unique;
$casesDispatchedToTheToQueue = $global:casesDispatchedToTheToQueue | Sort-Object | Get-Unique;
$casesDispatchedToTheFromQueue = $global:casesDispatchedToTheFromQueue | Sort-Object | Get-Unique;
$casesDispatchedToTheFromQueueAndThenLaterToTheToQueue = $global:casesDispatchedToTheFromQueueAndThenLaterToTheToQueue | Sort-Object | Get-Unique;
$casesDispatchedToBothQueues = Compare-object -ReferenceObject $casesDispatchedToTheFromQueue -DifferenceObject $casesDispatchedToTheToQueue -IncludeEqual -ExcludeDifferent;

write-log "Number of cases forwarded to the $toQueue queue : $($casesForwardedTo.count) "
write-log "Number of cases forwarded to the $toQueue queue from the $fromQueue queue: $($casesForwardedFromTo.count) "
write-log "Number of cases dispatched to the $toQueue queue : $($casesDispatchedToTheToQueue.count) "
write-log "Number of cases dispatched to the $fromQueue queue : $($casesDispatchedToTheFromQueue.count) "
write-log "Number of cases dispatched to both queues ($fromQueue and $toQueue) : $($casesDispatchedToBothQueues.count) "
write-log "Number of cases dispatched to the $fromQueue queue and then later dispatched to the $toQueue queue) : $($casesDispatchedToTheFromQueueAndThenLaterToTheToQueue.count) "

write-log "------------------------" -Level "Info";

if ($writeData){
	Write-Output "------------------------" 
	Write-Output "cases forwarded to the $toQueue queue " 
	Write-Output $casesForwardedTo;

	Write-Output "------------------------" 
	Write-Output "cases forwarded to the $toQueue queue from the $fromQueue queue" 
	Write-Output $casesForwardedFromTo;

	Write-Output "------------------------ " 
	Write-Output "cases dispatched to the $toQueue queue" 
	Write-Output $casesDispatchedToTheToQueue;


	Write-Output "------------------------" 
	Write-Output "cases dispatched to the $fromQueue queue " 
	Write-Output $casesDispatchedToTheFromQueue;


	Write-Output "------------------------" 
	Write-Output "casesDispatchedToTheFromQueue ($fromQueue) AndThenLaterToTheToQueue ($toQueue)"
	Write-Output $casesDispatchedToTheFromQueueAndThenLaterToTheToQueue;

	# Hint: keep this as the last output. The output of compare-object can be a bit odd at times.
	Write-Output "------------------------" 
	Write-Output "casesDispatchedToBothQueues  ($toQueue and $fromQueue)" 
	Write-Output $casesDispatchedToBothQueues;

}

