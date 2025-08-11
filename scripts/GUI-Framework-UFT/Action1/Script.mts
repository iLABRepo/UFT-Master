'*******************************************************************************************
'Test Name 				: Search Google
'Automated by 			: Bongs Lushaba
'Reviewd by 			: 
'Date					: 06/05/2018
'******************************************************************************************
If DataTable.GetCurrentRow = 1 Then
	'CreateFolder
CreateFolder(sResults)

CreateExcelReport sTestReport,"iLAB Framework"
DataTable.Import(strTestData)
End If
If DataTable.GetCurrentRow <> 1 Then
	 AppendToTextFile sTestReport,""
End If
 AppendToTextFile sTestReport,"scenario: "& DataTable.GetCurrentRow&";Test Steps;Status;Results"

Call SearchGoogle()


