On Error Resume Next
Set wordApp = CreateObject("Word.Application")
InpFolderPath = InputBox("Enter Input Folder Path")
OutputFolderPath = InputBox("Enter output folder path")
Set oFSO = CreateObject("Scripting.FileSystemObject")
wordApp.Visible = True
For Each oFile In oFSO.GetFolder(InpFolderPath).Files
  FileExtension = Ucase(oFSO.GetExtensionName(oFile.Name))	
  If FileExtension = "DOC" Or FileExtension = "DOCX" Or FileExtension = "DOCM" or FileExtension = "DOTM" Then
   	wordApp.Documents.Open(oFile.Path)
	wordApp.ActiveDocument.SaveAs(OutputFolderPath & "\" & oFile.Name)
	wordApp.ActiveDocument.Close
  End if
Next
wordApp.Quit

