On Error Resume Next

Set FSO1 = CreateObject("Scripting.FileSystemObject")
'tempFileName = "C:\Users\labuser\Desktop\Logging.txt"

tempFileName = "C:\MyData\Logging.txt"

Set oLog = FSO1.OpenTextFile(tempFileName, 2, True)

Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = "C:\Users\tarastog\Desktop\Office2016Files"

'objStartFolder = "C:\MyData\bugfiles"

oLog.WriteLine(objStartFolder)

Set wordApp = CreateObject("Word.Application")
Set pptApp = CreateObject("PowerPoint.Application")
Set xlApp = CreateObject("Excel.Application")

wordApp.Visible = True
pptApp.Visible = True
xlApp.Visible = True

Set objFolder = objFSO.GetFolder(objStartFolder)

oLog.WriteLine(objFolder)

Set colFiles = objFolder.Files
RecurseFiles colFiles,objFolder

ShowSubfolders objFSO.GetFolder(objStartFolder)

Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
		oLog.WriteLine("In " & Subfolder)
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        RecurseFiles colFiles,objFolder
		ShowSubFolders Subfolder
    Next
End Sub

Sub RecurseFiles(colFiles, objFolder)
	For Each oFile In colFiles
	  oLog.WriteLine(oFile)
	  FileExtension = Ucase(objFSO.GetExtensionName(oFile.Name))	
		If FileExtension = "DOC" Or FileExtension = "DOCX" Or FileExtension = "DOCM" or FileExtension = "DOTM" Then
			oLog.WriteLine("Processing Word file...")
			wordApp.Documents.Open(oFile.Path)
			wordApp.DisplayAlerts = False
			'wordApp.EnableEvents = False
			wordApp.ActiveDocument.Convert
			wordApp.ActiveDocument.SaveAs(objFolder & "\" & oFile.Name)
			wordApp.ActiveDocument.Close
		End If
	    If FileExtension = "PPT" Or FileExtension = "PPTX" Or FileExtension = "PPTM" or FileExtension = "POTM" Then
			oLog.WriteLine("Processing PowerPoint file...")
			pptApp.Presentations.Open(oFile.Path)
			pptApp.DisplayAlerts = False
			'pptApp.EnableEvents = False
			pptApp.ActivePresentation.Convert
			pptApp.ActivePresentation.SaveAs(objFolder & "\" & oFile.Name)
			pptApp.ActivePresentation.Close	
		End If
		If FileExtension = "XLS" Or FileExtension = "XLSX" Or FileExtension = "XLSM" or FileExtension = "XLTM" Then
			oLog.WriteLine("Processing Excel file...")
			xlApp.Workbooks.Open(oFile.Path)
			xlApp.DisplayAlerts = False
			'xlApp.EnableEvents = False
			xlApp.ActiveWorkbook.Convert
			xlApp.ActiveWorkbook.SaveAs(objFolder & "\" & oFile.Name)
			xlApp.ActiveWorkbook.Close	
		End If
	Next
End Sub

wordApp.Quit
xlApp.Quit
pptApp.Quit