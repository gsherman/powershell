$args_ = , $args_
$message=$args;

<#
Notes, testing v2 API:

* If the employee ID is invalid, then it throws a 500 error. WTF?
* If the casing of the key of a list value is not corret, it throws a 400 error. e.g. "medium" fails, but "Medium" succeeds.
* 
#>

$inputFile = "data.xlsx";
$username = "dovetail-api";
$password="letmein";
$url="http://localhost/api/v2/cases";
$debug = $true;
$showResponse = $false;

function CreateCase($row){

	$hash = @{  
		title = $row.title
		employee = $row.employeeID.toString();
		notes = $row.notes
		concerning = $row.concerningEmployeeId.toString();
		severity = $row.severity
		priority = $row.priority
	}

	$json = $hash | convertto-json;

	$credPair = "$($username):$($password)" 
	$encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($credPair)) 
	$requestHeaders = @{ Authorization = "Basic $encodedCredentials" } 

	write-host "----------------------------------"
	write-host "Creating Case for old case ID:" $row.ID;

	 try{
		$response = Invoke-RestMethod -Uri $url -Method Post -Body $json -ContentType "application/json" -Headers $requestHeaders 
		write-host "Success. Newly created case id:" $response.identifier;
		if ($showResponse -eq $true){ 
			write-host "--- RESPONSE ---"; 
			write-host $json; 
			write-host "--- END RESPONSE ------"; 
		}

	  }catch {
		$statusCode = $_.Exception.Response.StatusCode.value__
		$responseHeaders = $_.Exception.Response.Headers;
		
		write-host "Failure"
		write-host "Status Code: " $statusCode 
		if ($statusCode -eq 400){
			if ($responseHeaders.GetValues("X-Dovetail-BadListValue") ){
				Write-Host "Bad Value provided for List: "$responseHeaders.GetValues("X-Dovetail-BadListValue")
			}						
			if ($responseHeaders.GetValues("X-Dovetail-InvalidParameter") ){
				Write-Host "Invalid Parameter: "$responseHeaders.GetValues("X-Dovetail-InvalidParameter")
			}
			if ($responseHeaders.GetValues("X-Dovetail-MissingParameter") ){
				Write-Host "Missing Parameter: "$responseHeaders.GetValues("X-Dovetail-MissingParameter")
			}
		}
		if ($debug -eq $true){ 
			write-host "--- POSTED JSON ---"; 
			write-host $json; 
			write-host "--- END JSON ------"; 
		}
	 } # end catch block

} #end function


Install-module ImportExcel

write-host "Reading data from excel file:"  $inputFile;
$excel = Import-Excel -Path $inputFile;
write-host "Successfully read data from excel file. Number of rows: " $excel.count

foreach ($excelRow in $excel){
	CreateCase ($excelRow);
}
