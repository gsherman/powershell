# GetSolutionAttachmentDetails

Find all of the solution attachments

### Input - solution query download
Run a query within the Agent app for all solutions
Download to a file (Excel)
Rename to `Solutions.xslx`
Copy to same directory as the GetSolutionAttachmentDetails.ps1 script

### Setup
Edit the GetSolutionAttachmentDetails.ps1 file, setting:
* username
* password
* getSolutionUrl

To run, from a PowerShell prompt: `.\GetSolutionAttachmentDetails.ps1 .\solutions.xlsx`

Output will be `attachments.csv`, which is a CSV file of all solution attachments

Logs are written to the `\logs` directory

