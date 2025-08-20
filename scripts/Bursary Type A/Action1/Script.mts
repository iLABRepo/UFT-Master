Dim excelPath, fso, scriptsPath, path
Dim data
Set fso =  CreateObject("Scripting.FileSystemObject")
scriptsPath = fso.GetParentFolderName(Environment("TestDir"))
path = fso.GetParentFolderName(scriptsPath)

excelPath = path & "\data\Wits Staff Bursary(1).xlsx"

DataTable.Import excelPath

Call LaunchBrowser(DataTable.Value("browser"), DataTable.Value("url"))

Call Login(DataTable.Value("Username"), DataTable.Value("Password")) @@ script infofile_;_ZIP::ssf193.xml_;_

Call StartHere()

Call BursaryTypeLink(DataTable.Value("BursaryType")) @@ script infofile_;_ZIP::ssf202.xml_;_

Call GettingStarted()

Call AgreeToTermsAndConditions() @@ script infofile_;_ZIP::ssf211.xml_;_

Call HighestQualificationAndHistoricalRecord(DataTable.Value("QualificationString"), DataTable.Value("PreviousQualificationName"), DataTable.Value("educationInstitution"), DataTable.Value("witsBursaryRecievedBefore"), DataTable.Value("yearOfQualification"), DataTable.Value("yearOfRegistration"), DataTable.Value("Full_PartTime"), DataTable.Value("BursaryType"))
 @@ script infofile_;_ZIP::ssf10.xml_;_
Call EnrollmentDetails(DataTable.Value("SearchString"), DataTable.Value("ProgramName"),DataTable.Value("YearOfStudy"), DataTable.Value("TotalDuration"), DataTable.Value("PartTimeOrFullTime"), DataTable.Value("NQFLevel")) @@ script infofile_;_ZIP::ssf233.xml_;_

Call SupportingDocumentsForBursary(path & DataTable.Value("UploadedFilePath"), path & DataTable.Value("UploadedFilePathTxt"), path & DataTable.Value("UploadedFilePathDocx"))

Call ClickSubmit()

Call ClickLogout()


























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




