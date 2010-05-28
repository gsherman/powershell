$strPath = "c:\temp\opencases.xls";

. .\DovetailCommonFunctions.ps1

$ClarifyApplication = create-clarify-application; 
$ClarifySession = create-clarify-session $ClarifyApplication; 

$dataSet = new-object FChoice.Foundation.Clarify.ClarifyDataSet($ClarifySession)
$caseGeneric = $dataSet.CreateGeneric("extactcase")
$caseGeneric.AppendFilter("status", "Equals", "Working")
$caseGeneric.AppendFilter("condition", "Equals", "Open")
$caseGeneric.Query()

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$workbook = $excel.Workbooks.add()
$sheet = $workbook.worksheets.Item(1)
$x = 1;

foreach( $case in $caseGeneric.Rows){
 $sheet.cells.item($x, 1) = $case["id_number"]
 $sheet.cells.item($x,2) = $case["title"]
 $x++
}

if(Test-Path $strPath)
  { 
   Remove-Item $strPath
   $excel.ActiveWorkbook.SaveAs($strPath)
  }
else
  {
   $excel.ActiveWorkbook.SaveAs($strPath)
  }
     