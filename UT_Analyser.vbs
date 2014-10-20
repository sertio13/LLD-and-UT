'Script to Analyze Unit Test Cases


'Intialize File and Shell Objects
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.shell")

'Set Path where the Excel files are located
Set objFolder = objFSO.GetFolder("C:\Users\\Desktop\Scripts\UT")
Set colFiles = objFolder.Files

'Set Path of Output Files
Set ApprovedUT = objFSO.CreateTextFile("C:\Users\\Desktop\Scripts\UT_OK.txt", True)
Set NeedtoActUT = objFSO.CreateTextFile("C:\Users\\Desktop\Scripts\UT_NOT_OK.txt", True)

'To access every excel file in the folder
For Each objFile in colFiles
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open _
    (objFile)

'Activate Sheet 2 of the Excel Sheet 
Set objWorksheet = objExcel.Worksheets(2)
	objWorksheet.Activate
	
	On Error resume next 'Prevents the script from crashing when  it encounters an error
	'Wscript.Echo objExcel.Cells(3,10).Value
	'Wscript.Echo (objFile)
	
'Read Value from excel sheet to check for Approved or Baselined UT's
	If objExcel.Cells(3,10).Value = "Approved" or "Baselined" Then
		ApprovedUT.WriteLine("")
		ApprovedUT.Write("UT Status : ")
		ApprovedUT.WriteLine(objExcel.Cells(3,10).Value)
		ApprovedUT.Write("File Name : ") 
		ApprovedUT.WriteLine(objFile)
		
	
		
		
		
		'ApprovedUT.Write("UT ID : ") 
		'ApprovedUT.WriteLine(objExcel.Cells(2,14).Value)
			'If NOT objFile = (objExcel.Cells(2,14).Value) Then
			'ApprovedUT.Writeline("Action : UT needs to be renamed")
			'End If
			
		ApprovedUT.Write("UT Prepared By : ") 
		ApprovedUT.WriteLine(objExcel.Cells(2,8).Value)
		

		
		
		
	Else
				NeedtoActUT.WriteLine("")
				NeedtoActUT.Writeline(objFile)
				NeedtoActUT.Write("UT Status : ")
				NeedtoActUT.WriteLine(objExcel.Cells(3,10).Value)
				NeedtoActUT.Write("UT Prepared By : ") 
				NeedtoActUT.WriteLine(objExcel.Cells(2,8).Value)
		
		
		
	End If	
		
	

objExcel.Quit
 
Next
'Call powershell script to send email 
'objShell.run("powershell -noexit -file C:\Users\\Desktop\Scripts\PS\outlook_email.ps1")