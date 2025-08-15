Dim excelPath, fso, scriptsPath, path
Dim data
Set fso =  CreateObject("Scripting.FileSystemObject")
scriptsPath = fso.GetParentFolderName(Environment("TestDir"))
path = fso.GetParentFolderName(scriptsPath)

excelPath = path & "\data\Wits Staff Bursary(1).xlsx"

DataTable.Import excelPath

DataTable.SetCurrentRow 2

Login DataTable.Value("Username"), DataTable.Value("Password") 

ClickStartHere

ClickBursaryTypeLink DataTable.Value("BursaryType")

ClickNext

AgreeToTermsAndConditions

DependantInformation DataTable.Value("witsStudentNumber"), DataTable.Value("title"),DataTable.Value("natureOfRelationship"), DataTable.Value("contactNumber"), DataTable.Value("disabilities"), DataTable.Value("grossIncome"), DataTable.Value("estimatedGrossIncome"), DataTable.Value("taxNumber")

HighestQualificationAndHistoricalRecord DataTable.Value("QualificationString"), DataTable.Value("PreviousQualificationName")
 
EnrollmentDetails DataTable.Value("SearchString"), DataTable.Value("ProgramName"), DataTable.Value("YearOfStudy"), DataTable.Value("TotalDuration"), DataTable.Value("PartTimeOrFullTime"), DataTable.Value("NQFLevel") @@ script infofile_;_ZIP::ssf12.xml_;_
 @@ script infofile_;_ZIP::ssf4.xml_;_

SupportingDocumentsForBursary path & DataTable.Value("UploadedFilePath")

ClickSubmit DataTable.Value("SubmissionHeading")

ClickLogout
