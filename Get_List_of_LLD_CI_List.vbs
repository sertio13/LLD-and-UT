Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.shell")
strFolder = InputBox("Enter the path where the LLD's are located:")
Set objFolder = objFSO.GetFolder(strFolder)
Set colFiles = objFolder.Files
Set LLDList = objFSO.CreateTextFile("C:\Users\\Desktop\Scripts\CI_List_LLD.csv", True)
For Each objFile in colFiles
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open _
    (objFile)

Set objWorksheet = objExcel.Worksheets("LLD")
	objWorksheet.Activate
	On Error resume next
	
		LLDList.WriteLine("")
				LLDList.Write("File Name :") 
		LLDList.WriteLine(objFile)
		LLDList.Write("LLD ID : ") 
		LLDList.WriteLine(objExcel.Cells(2,14).Value)
			

objExcel.Quit
 
Next
Wscript.Echo " Completed! Find the list at C:\Users\\Desktop\Scripts\LLD_List.csv"