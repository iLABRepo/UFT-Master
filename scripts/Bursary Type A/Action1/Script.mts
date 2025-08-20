Dim excelPath, fso, scriptsPath, path
Dim data
Set fso =  CreateObject("Scripting.FileSystemObject")
scriptsPath = fso.GetParentFolderName(Environment("TestDir"))
path = fso.GetParentFolderName(scriptsPath)
 
excelPath = path & "\data\Wits Staff Bursary(1).xlsx"
 
DataTable.Import excelPath
 
 
Call Login(DataTable.Value("Username"), DataTable.Value("Password"))
 
Call StartHere()
 
Call BursaryTypeLink(DataTable.Value("BursaryType"))
Call GettingStarted()
Call AgreeToTermsAndConditions()
 
Call HighestQualificationAndHistoricalRecord(DataTable.Value("QualificationString"), DataTable.Value("PreviousQualificationName"))
 
Call EnrollmentDetails(DataTable.Value("SearchString"), DataTable.Value("ProgramName"),DataTable.Value("YearOfStudy"), DataTable.Value("TotalDuration"), DataTable.Value("PartTimeOrFullTime"), DataTable.Value("NQFLevel"))
 
Call SupportingDocumentsForBursary(path & DataTable.Value("UploadedFilePath"))
 
Call ClickSubmit()
 
Call ClickLogout()

'Line Manager accept test




Dim excelPath, fso, scriptsPath, path
Dim data
Set fso =  CreateObject("Scripting.FileSystemObject")
scriptsPath = fso.GetParentFolderName(Environment("TestDir"))
path = fso.GetParentFolderName(scriptsPath)
 
excelPath = path & "\data\Wits Staff Bursary(1).xlsx"
 
DataTable.Import excelPath

'Call LineManagerApproval()
Call HRApproval()

 @@ script infofile_;_ZIP::ssf24.xml_;_





