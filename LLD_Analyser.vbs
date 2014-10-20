Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.shell")
strFolder = InputBox("Enter the path where the LLD's are located:")
Set objFolder = objFSO.GetFolder(strFolder)
Set colFiles = objFolder.Files
Set ApprovedLLD = objFSO.CreateTextFile("C:\Users\minda\Desktop\Scripts\LLD_OK.txt", True)
Set NeedtoActLLD = objFSO.CreateTextFile("C:\Users\minda\Desktop\Scripts\LLD_NOT_OK.txt", True)
For Each objFile in colFiles
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open _
    (objFile)

Set objWorksheet = objExcel.Worksheets("LLD")
	objWorksheet.Activate
	On Error resume next
	'Wscript.Echo " " & objExcel.Cells(4,14).Value
	'Wscript.Echo (objFile)
	If objExcel.Cells(4,14).Value = "Approved" Then
		ApprovedLLD.WriteLine("")
		ApprovedLLD.Write("LLD Status : ")
		ApprovedLLD.WriteLine(objExcel.Cells(4,14).Value)
		ApprovedLLD.Write("File Name :") 
		ApprovedLLD.WriteLine(objFile)
		ApprovedLLD.Write("LLD ID : ") 
		ApprovedLLD.WriteLine(objExcel.Cells(2,14).Value)
			'If NOT objFile = (objExcel.Cells(2,14).Value) Then
			'ApprovedLLD.Writeline("Action : LLD needs to be renamed")
			'End If
			
		ApprovedLLD.Write("LLD Prepared By : ") 
		ApprovedLLD.WriteLine(objExcel.Cells(3,12).Value)
		
		
	Else
				NeedtoActLLD.WriteLine("")
				NeedtoActLLD.Writeline(objFile)
				NeedtoActLLD.Write("LLD Status : ")
				NeedtoActLLD.WriteLine(objExcel.Cells(4,14).Value)
				NeedtoActLLD.Write("LLD Prepared By : ") 
				NeedtoActLLD.WriteLine(objExcel.Cells(3,12).Value)
		
		
		
	End If	
		
	

objExcel.Quit
 
Next
Wscript.Echo "Completed! Open LLD_OK.txt or LLD_NOT_OK.txt to get the details" 
'objShell.run("powershell -noexit -file C:\Users\minda\Desktop\Scripts\PS\outlook_email.ps1")