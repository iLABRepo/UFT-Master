SystemUtil.Run "msedge.exe", "https://qa.intranetapps.wits.ac.za/was/applogin"
Dim excelPath, fso, scriptsPath, path
Dim data
Set fso =  CreateObject("Scripting.FileSystemObject")
scriptsPath = fso.GetParentFolderName(Environment("TestDir"))
path = fso.GetParentFolderName(scriptsPath)
 
excelPath = path & "\data\Wits Staff Bursary(1).xlsx"
 
DataTable.Import excelPath
 
DataTable.SetCurrentRow(2)
 
Call Login(DataTable.Value("Username"), DataTable.Value("Password"))
 
Call StartHere()
 
Call BursaryTypeLink(DataTable.Value("BursaryType"))

Call GettingStarted()

Call AgreeToTermsAndConditions()

Call DependantInformation(DataTable.Value("witsStudentNumber"), DataTable.Value("title"),DataTable.Value("natureOfRelationship"), DataTable.Value("contactNumber"), DataTable.Value("disabilities"), DataTable.Value("grossIncome"), DataTable.Value("estimatedGrossIncome"), DataTable.Value("taxNumber"))
 



















































































































































Call HighestQualificationAndHistoricalRecord(DataTable.Value("QualificationString"), DataTable.Value("PreviousQualificationName"), DataTable.Value("educationInstitution"), DataTable.Value("witsBursaryRecievedBefore"), DataTable.Value("yearOfQualification"), DataTable.Value("yearOfRegistration"), DataTable.Value("Full_PartTime"))
 
Call EnrollmentDetails(DataTable.Value("SearchString"), DataTable.Value("ProgramName"),DataTable.Value("YearOfStudy"), DataTable.Value("TotalDuration"), DataTable.Value("PartTimeOrFullTime"), DataTable.Value("NQFLevel"))
 
Call SupportingDocumentsForBursary(path & DataTable.Value("UploadedFilePath"))
 
Call ClickSubmit()
 
Call ClickLogout()
