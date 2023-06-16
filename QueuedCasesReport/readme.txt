Run the report: queued cases since last year
export to Excel (xls)
remove the extra header rows; leave the row that contains the header column names (CaseID, caseCreated,Action, etc.)
from powershell, run the analysiss script, passing in the excel file, FromQueue, and ToQueue 
e.g. 
.\analysis.ps1 -inputfile cases.xlsx "General HR" Benefits

At the moment, there's a formula in the Exago report that translates the queueID (GUID) to a queueName. 
This works for customers with small numbers of queues.
If there's a lot of queues, we may wish to:
- have the report only output queue IDs, and not do the queueID --> queueName mapping 
-  modify the powershell script, and have teh script handle the queueID --> queueName mapping 

