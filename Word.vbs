On Error Resume Next

Set FSO1 = CreateObject("Scripting.FileSystemObject")
tempFileName = "C:\MyData\Logging.txt"
Set oLog = FSO1.OpenTextFile(tempFileName, 2, True)

Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = "C:\Users\tarastog\Desktop\Office2016Files"

oLog.WriteLine(objStartFolder)

Set wordApp = CreateObject("Word.Application")

wordApp.Visible = True

Set objFolder = objFSO.GetFolder(objStartFolder)

oLog.WriteLine(objFolder)

Set colFiles = objFolder.Files

For Each oFile In colFiles
  oLog.WriteLine(oFile)
  FileExtension = Ucase(objFSO.GetExtensionName(oFile.Name))	
  If FileExtension = "DOC" Or FileExtension = "DOCX" Or FileExtension = "DOCM" or FileExtension = "DOTM" Then
	oLog.WriteLine("Processing...")
   	wordApp.Documents.Open(oFile.Path)
	wordApp.ActiveDocument.SaveAs(objFolder & "\" & oFile.Name)
	wordApp.ActiveDocument.Close
  End if
Next

ShowSubfolders objFSO.GetFolder(objStartFolder)

Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
		oLog.WriteLine("In " & Subfolder)
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        For Each oFile In colFiles
		  oLog.WriteLine(oFile)
		  FileExtension = Ucase(objFSO.GetExtensionName(oFile.Name))	
		  If FileExtension = "DOC" Or FileExtension = "DOCX" Or FileExtension = "DOCM" or FileExtension = "DOTM" Then
			oLog.WriteLine("Processing...")
			wordApp.Documents.Open(oFile.Path)
			wordApp.ActiveDocument.SaveAs(objFolder & "\" & oFile.Name)
			wordApp.ActiveDocument.Close
		  End if
		Next
		ShowSubFolders Subfolder
    Next
End Sub

wordApp.Quit