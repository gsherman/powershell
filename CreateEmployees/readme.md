# Import Employees 

Import employees from ax Excel sheet into a Dovetail system.
Uses the Dovetail create-employee API.

### Setup
Edit the file, setting the 
* Global Variables (username, password, url)
* Array of CustomFieldNames (these are the columns in the Excel file that you want to be custom fields)
* $global:logLevel (set to DEBUG level 4 during initial testing)

To run, from a PowerShell prompt: `.\ImportEmployees.ps1 .\employees.xlsx`

Logs are written to the `\logs` directory

