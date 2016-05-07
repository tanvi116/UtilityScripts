set pdfMaker = CreateObject("PDFMakerApp.PDFMakerDriver.1")
set application = pdfMaker.launchApplication(3) ' PowerPoint
'set application = pdfMaker.launchApplication(1) ' Word
Wscript.echo "Application started."
application.AutomationSecurity = 1
application.DisplayAlerts = 1
Wscript.echo "Starting the conversions."

For i = 0 To 500
	set conversionSettings = CreateObject("PDFMakerAPI.ConversionSettings")
	Wscript.Echo "1"
	status = conversionSettings.LoadSettingsFromFile("Standard")
	Wscript.Echo "2"
	status = conversionSettings.SetConversionParameters(-2042585081)
	Wscript.Echo "3"
	conversionSettings.SetInvokedFromFeat(true)
	Wscript.Echo "4"

'	application.Visible = false
	wscript.Echo "5"
	set dispatchDocument = pdfMaker.OpenDocument("C:\standalone_collateral\HelloWorld.pptx")
'	set dispatchDocument = pdfMaker.OpenDocument("C:\standalone_collateral\HelloWorld.docx")
	Wscript.Echo "6"

	strPdfFilePath = "C:\standalone_collateral\HelloWorld" & i & ".pdf"
	Wscript.Echo "7"
	result = pdfMaker.ExecPDFMaker(strPdfFilePath, conversionSettings, true, false)
	Wscript.Echo "8"

	If Eval("result = 0") Then
		Wscript.Echo "Conversion#" & i & " Successful"
	Else 
		Wscript.Echo "Conversion Failed"
	End If
	pdfMaker.CloseDocument()
	Wscript.Echo "Closed document"

Next

pdfMaker.closeApplication
Wscript.echo "Application closed"
