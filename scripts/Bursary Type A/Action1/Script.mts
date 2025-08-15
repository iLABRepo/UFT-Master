Dim excelPath, fso, scriptsPath, path
Dim data
Set fso =  CreateObject("Scripting.FileSystemObject")
scriptsPath = fso.GetParentFolderName(Environment("TestDir"))
path = fso.GetParentFolderName(scriptsPath)

excelPath = path & "\data\Wits Staff Bursary(1).xlsx"

DataTable.Import excelPath


Call Login(DataTable.Value("Username"), DataTable.Value("Password"))

Call ClickStartHere()

Call ClickBursaryTypeLink(DataTable.Value("BursaryType"))
Call ClickNext()
Call AgreeToTermsAndConditions()

Call HighestQualificationAndHistoricalRecord(DataTable.Value("QualificationString"), DataTable.Value("PreviousQualificationName"))
 @@ script infofile_;_ZIP::ssf10.xml_;_
Call EnrollmentDetails(DataTable.Value("SearchString"), DataTable.Value("ProgramName"),DataTable.Value("YearOfStudy"), DataTable.Value("TotalDuration"), DataTable.Value("PartTimeOrFullTime"), DataTable.Value("NQFLevel"))

Call SupportingDocumentsForBursary(path & DataTable.Value("UploadedFilePath"))

Call ClickSubmit(DataTable.Value("SubmissionHeading"))

ClickLogout


























'Line Manager accept test
'DataTable.Import excelPath

'DataTable.SetCurrentRow 1

'LineManagerUsername
'Login DataTable.Value("LineManagerUsername"), DataTable.Value("Password")

'ClickActionsRequired

'ClickBursaryTypeLink DataTable.Value("BursaryType")
'ClickOnTheFirstSubmission

'ClickApprove

'SubmitRequest DataTable.Value("HRApprover")




