param ([Parameter(Mandatory)]$inputFile, $dryRun=$true)

<#
Calling syntax:
	.\importemployees.ps1 input.xlsx 

Outputs: 
 	* log file, in the \logs directory
#>


# Global Variables
$username = "dovetail-api";
$password="letmein";
$url="http://localhost/api/v5/employees";

# Custom Fields
$customFieldNames =  @();
$customFieldNames+= "Flight Risk";
$customFieldNames+= "Manager";
$customFieldNames+= "Salary";
$customFieldNames+= "Performance Indicator";
$customFieldNames+= "emojis 💥❤️✔️";

# For Testing Use Only
# Append this to the end of certain fields (employeeID, username), to allow for uniqueness
# set to empty string for normal use
$testingFieldAppendix="";

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
	$href= $responseObject.href;
	$warningsArray= $responseObject.warnings;
	$errorsArray= $responseObject.errors;

	if ($status -eq 201){
		$href= $responseObject.href;

		write-log -Message "Success 201. Newly created employee: $($href)" -Level "Info";

		$global:pass++;

	 }else{
	 	write-log -Message "Create Employee Failed on row $rowIndex" -Level "Error"
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

	$employee = @{}

	if ($row.FirstName) { $employee.firstName = $row.FirstName.toString().Trim(); }
	if ($row.LastName) { $employee.lastName = $row.LastName.toString().Trim(); }
	if ($row.EmployeeID) { $employee.employeeID = $row.EmployeeID.toString().Trim() + $testingFieldAppendix; }
	if ($row.HRISId) { $employee.HRISId = $row.HRISId.toString().Trim(); }
	if ($row.EmployeeTitle) { $employee.title = $row.EmployeeTitle.toString().Trim(); }
	if ($row.PreferredFirstName) { $employee.preferredFirstName = $row.PreferredFirstName.toString().Trim(); }
	if ($row.MiddleName) { $employee.middleName = $row.MiddleName.toString().Trim(); }
	if ($row.Status) { $employee.status = $row.Status.toString().Trim(); }
	if ($row.Suffix) { $employee.suffix = $row.Suffix.toString().Trim(); }
	if ($row.Pronouns) { $employee.pronouns = $row.Pronouns.toString().Trim(); }
	if ($row.SiteID) { $employee.primarySiteId = $row.SiteID.toString().Trim(); }
	if ($row.isContingent) { $employee.isContingent = $row.isContingent.toString().Trim(); }
	if ($row.JobTitle) { $employee.jobTitle = $row.JobTitle.toString().Trim(); }
	if ($row.PortalSSOUsername) { $employee.username = $row.PortalSSOUsername.toString().Trim()  + $testingFieldAppendix; }
	if ($row.PortalAccess) { $employee.loginStatus = $row.PortalAccess.toString().Trim(); }
	if ($row.Culture) { $employee.preferredCulture = $row.Culture.toString().Trim(); }
	if ($row.AvatarUrl) { $employee.avatarUrl = $row.AvatarUrl.toString().Trim(); }
	
	if ($row.PrimaryEmail){ 
		$email =  @{};
		$email.emailType = "Work";
		$email.email = $row.PrimaryEmail.toString().Trim();
		$employee.primaryEmailAddress = $email;
	}

	if ($row.PrimaryPhone){ 
		$phone =  @{};
		$phone.phoneType = "Work";
		$phone.phoneNumber = $row.PrimaryPhone.toString().Trim();
		$employee.primaryPhoneNumber = $phone;
	}

	# Additional phone numbers
	$mobilePhone =  @{};
	$mobilePhone.phoneType = "Mobile";
	$mobilePhone.phoneNumber = $row.CellPhoneNumber.toString().Trim();

	$otherPhone =  @{};
	$otherPhone.phoneType = "Other";
	$otherPhone.phoneNumber = $row.CellPhoneNumber.toString().Trim();

	$phoneArray =  @();
	$phoneArray+= $mobilePhone;
	$phoneArray+= $otherPhone;
	$employee.phoneNumbers = $phoneArray;

	# Additional emails
	$homeEmail =  @{};
	$homeEmail.emailType = "Home";
	$homeEmail.email = $row.PrimaryEmail.toString().Trim();

	$alternateEmail =  @{};
	$alternateEmail.emailType = "Other";
	$alternateEmail.email = $row.PrimaryEmail.toString().Trim();

	$emailArray =  @();
	$emailArray+= $homeEmail;
	$emailArray+= $alternateEmail;
	$employee.emailAddresses = $emailArray;
		
	# Employee Filters
	if ($row.EmployeeFilters) { 
		$employeeFiltersArray = @($row.EmployeeFilters.toString().Split(",").Trim() );		
		$employee.employeeFilters = $employeeFiltersArray;		
	}			

	# Portal Filters
	if ($row.PortalFilters) { 
		$portalFiltersArray = @($row.PortalFilters.toString().Split(",").Trim() );		
		$employee.portalFilters = $portalFiltersArray;		
	}	

	# Tags
	if ($row.Tags) { 
		$tagsArray = @($row.Tags.toString().Split(",").Trim() );		
		$employee.tags = $tagsArray;		
	}


	$customFields =  @();
	$customFieldNames.ForEach(
		{ 
			if ($row.$_){
				$customField =  @{};
				$customField.value = $row.$_.toString().Trim();
				$customField.key = $_;
				$customFields+=$customField;

			}
		}
	)
	$employee.customFields = $customFields

	# convert to json
	$jsonRequestBody = $employee | convertto-json -Depth 5;
	$jsonRequestBody
}


##################
# Create Employee
##################
function CreateEmployee{
 param( $row, [int]$rowIndex )

	write-log " "  -Level "Info";
	write-log "--------------------------"  -Level "Info";
	write-log "Processing row $rowIndex" -Level "Info"; 

	$jsonRequestBody = BuildRequestBody $row $rowIndex;

	write-log -Message $jsonRequestBody.ToString() -Level "Debug"; 

	# testing
	# return;

	$credPair = "$($username):$($password)" 
	$encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($credPair)) 
	$requestHeaders = @{ 
		Authorization = "Basic $encodedCredentials" 
		Accept = "application/json"
	} 

	try{
		#$response = Invoke-WebRequest -Uri $url -Method Post -Body $jsonRequestBody -ContentType "application/json" -Headers $requestHeaders 
		$response = Invoke-WebRequest -Uri $url -Method Post -Body ([System.Text.Encoding]::UTF8.GetBytes($jsonRequestBody)) -ContentType "application/json" -Headers $requestHeaders 
		

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
	CreateEmployee $excelRow $rowCounter;
}

 write-log "------------------------" -Level "Info";
 write-log "Pass: $($global:pass) of $($excel.count)" -Level "Info";
 write-log "Fail: $($global:fail) of $($excel.count)" -Level "Info";
