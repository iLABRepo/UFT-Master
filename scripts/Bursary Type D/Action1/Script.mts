Dim excelPath, fso, scriptsPath, path
Dim bursaryType
Dim data
bursaryType = "Bursary Type D: Staff dependant studying at Wits"
Set fso =  CreateObject("Scripting.FileSystemObject")
scriptsPath = fso.GetParentFolderName(Environment("TestDir"))
path = fso.GetParentFolderName(scriptsPath)
 
excelPath = path & "\data\Wits Staff Bursary(1).xlsx"
 
DataTable.Import excelPath

Call LaunchBrowser(DataTable.Value("browser"), DataTable.Value("url"))
 
Call Login(DataTable.Value("Username"), DataTable.Value("Password"))
 
Call ClickStartHere()
 
Call BursaryTypeLink(bursaryType)

Call GettingStarted()

Call AgreeToTermsAndConditions()

Call DependantInformation(DataTable.Value("witsStudentNumber"), DataTable.Value("title"),DataTable.Value("natureOfRelationship"), DataTable.Value("contactNumber"), DataTable.Value("disabilities"), DataTable.Value("grossIncome"), DataTable.Value("estimatedGrossIncome"), DataTable.Value("taxNumber"))


Call HighestQualificationAndHistoricalRecord(DataTable.Value("QualificationString"), DataTable.Value("PreviousQualificationName"), DataTable.Value("educationInstitution"), DataTable.Value("witsBursaryRecievedBefore"), DataTable.Value("yearOfQualification"), DataTable.Value("yearOfRegistration"), DataTable.Value("Full_PartTime"), bursaryType)
 
Call EnrollmentDetails(DataTable.Value("SearchString"), DataTable.Value("ProgramName"),DataTable.Value("YearOfStudy"), DataTable.Value("TotalDuration"), DataTable.Value("PartTimeOrFullTime"), DataTable.Value("NQFLevel"))
 
Call SupportingDocumentsForBursary(path & DataTable.Value("UploadedFilePath"), path & DataTable.Value("UploadedFilePathTxt"), path & DataTable.Value("UploadedFilePathDocx"))
 
Call ClickSubmit(bursaryType & " Submissions")
 
Call ClickLogout()

Call LineManagerApproval(bursaryType)

Call HRApproval(bursaryType)
