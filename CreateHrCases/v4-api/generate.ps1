# Generates an excel file of cases that can then be loaded with the ImportCases.ps1 script

# How many cases to generate
$numCases = 1;

# build the excel file name
$excelFileName = "$($numCases)-cases.xlsx";

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
		            Write-Host "$FormattedDate : $Message" -ForegroundColor Red;
		            $LevelText = 'ERROR:'
		            "$FormattedDate $LevelText $Message" | Out-File -FilePath $logFile -Append
	            }
            }
        'Warn' {
	        	if ($global:logLevel -gt 1){
		            Write-Warning "$FormattedDate : $Message"
		            $LevelText = 'WARNING:'
		            "$FormattedDate $LevelText $Message" | Out-File -FilePath $logFile -Append
	            }            
            }
        'Info' {
				if ($global:logLevel -gt 2){
		            Write-Host "$FormattedDate : $Message"
		            $LevelText = 'INFO:'
		            "$FormattedDate $LevelText $Message" | Out-File -FilePath $logFile -Append
	            }
            }
         'Debug'{
				if ($global:logLevel -gt 3){
					Write-Host "$FormattedDate : $Message"  -ForegroundColor Blue; 
					$LevelText = 'DEBUG:'
					"$FormattedDate $LevelText $Message" | Out-File -FilePath $logFile -Append
				}
         	}  

        } #end switch
        
} #end function


##################
# Get Random Date
##################
Function Get-RandomDate {
    [cmdletbinding()]
    param(
        [DateTime]
        $Min,

        [DateTime]
        $Max = [DateTime]::Now
    )

    Begin{
        If(!$Min -or !$Max){
            Write-Warning "Unable to parse entered string for Min or Max parameter. Proper example: `"06/23/1996 14:06:03.297`""
            Write-Warning "Time will default to midnight if omitted"
            Break
        }
    }

    Process{
        $randomTicks = Get-Random -Minimum $Min.Ticks -Maximum $Max.Ticks
        New-Object DateTime($randomTicks)
    }
}

##################
# Main
##################
If ((Get-Module -ListAvailable -Name "ImportExcel") -eq $null){
	Install-module ImportExcel;
}

If ((Get-Module -ListAvailable -Name "NameIT") -eq $null){
	Install-Module NameIT;
}

If ((Get-Module -ListAvailable -Name "LoremText") -eq $null){
	Install-Module LoremText;
}

 . .\LoremText.ps1

$ExcelParams = @{
        Path    = $excelFileName
        Show    = $false
        Verbose = $true
    }
Remove-Item -Path $ExcelParams.Path -Force -EA Ignore

$rows =  [System.Collections.ArrayList]@();

write-log "Starting loop for $($numCases)" -Level "Info";

for ($i=1; $i -le $numCases; $i++) {

	$dt = Get-RandomDate -Min "06/23/1996 14:06:03.297";
	$createDate = $dt.ToString("MM/dd/yy HH:mm:ss");
	$closeDate = $dt.AddHours(2);
	$closeDateString = $closeDate.ToString("MM/dd/yy HH:mm:ss");
	$notes = LoremText -Paragraphs 3;
	$title = LoremText -Paragraphs 1 -Sentences 1;

	$row = [PSCustomObject]@{
		ID	= 'old_'+$i
		Title = $title	
		Notes = $notes
		employeeId = 21118
		concerningEmployeeId= ''	
		condition= 'Closed'		
		closeNotes= 'These are closing notes'		
		closeResolution= 'QuickClose'	
		closeDate= $closeDateString	
		status= 'Researching'	
		caseType= 'Policies'		
		portalCaseType= 'GeneralHR'		
		severity= 'Low'		
		priority= 'High'		
		origin= 'Manual Entry'		
		availableInPortal= 'True'		
		sensitive= 'True'		
		queue	= ''	
		originatorUserName= 'systemservice'		
		ownerUserName= 'andrew'		
		createDate= $createDate		
		labels= 'Red, Silver'		
		createEvents= 'False'
		}
    $null = $rows.Add($row);
}

write-log "Finished loop for $($numCases)" -Level "Info";

write-log "starting write to excel" -Level "Info";

$rows | Export-Excel @ExcelParams -FreezeTopRow -AutoSize -MaxAutoSizeRows 3 -NoNumberConversion *; 

write-log "finished write to excel" -Level "Info";
