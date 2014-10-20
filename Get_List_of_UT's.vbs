Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.shell")
strFolder = InputBox("Enter the path where the UT's are located:")
Set objFolder = objFSO.GetFolder(strFolder)
Set colFiles = objFolder.Files
Set UList = objFSO.CreateTextFile("C:\Users\\Desktop\Scripts\UT_List.csv", True)
For Each objFile in colFiles
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open _
    (objFile)

Set objWorksheet = objExcel.Worksheets("UT")
	objWorksheet.Activate
	On Error resume next
	
		UList.WriteLine("")
				UList.Write("File Name :") 
		UList.WriteLine(objFile)
		
		UList.Write("UT ID : ") 
		UList.WriteLine(objExcel.Cells(2,7).Value)
		
		UList.Write("LLD ID : ") 
		UList.WriteLine(objExcel.Cells(2,4).Value)
	
	
		

objExcel.Quit
 
Next

Wscript.Echo "Completed! Find the list at C:\Users\\Desktop\Scripts\UT_List.csv"