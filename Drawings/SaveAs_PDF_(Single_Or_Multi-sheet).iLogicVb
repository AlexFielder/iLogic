Option Explicit On
'''Prints the selected sheets to multi-page pdf or single sheet pdf file(s).
''' 
Sub Main()
	Dim oDocument As DrawingDocument = ThisApplication.ActiveDocument
	Dim oSheets As Sheets
	Dim oSheet As Sheet
	Dim DrawingNumber As String = String.Empty
	Dim DrawingTitle As String = String.Empty
	Dim DrawingPrefix As String = String.Empty
	Dim oPar As UserParameter
	'Dim shtmaxnum as Integer = Convert.ToInt32(Parameter("Custom_End_Sheet"))
	'MessageBox.Show(shtmaxnum)
	Dim dwgdoc As DrawingDocument = ThisApplication.ActiveDocument
	
	For i = 1 To oDocument.Sheets.Count
		oDocument.Sheets(i).Activate
		oSheet = oDocument.Sheets(i)
		'MessageBox.Show(osheet.Name)
		Dim oTitleBlock As TitleBlock =oDocument.Sheets(i).TitleBlock
		'oTextBoxes=oTitleBlock.Definition.Sketch.TextBoxes
		For Each oTextBox as Inventor.TextBox In oTitleBlock.Definition.Sketch.TextBoxes
			' MessageBox.Show(otextbox.Text)
			If oTextBox.Text <> "" Then
				Select oTextBox.Text
					Case "<Drawing Number>":
						DrawingNumber = oTitleBlock.GetResultText(oTextBox)
					Case "Drawing Prefix": 
						DrawingPrefix = oTitleBlock.GetResultText(oTextBox)
					Case "TITLE":
						DrawingTitle = oTitleBlock.GetResultText(oTextBox)
				End Select
			End If
		Next
		If Not DrawingNumber = String.Empty And Not DrawingPrefix = String.Empty Then
			Logger.Debug("printing: " & oSheet.Name & " " & DrawingNumber & " " & DrawingPrefix)
			PrintSheet(oSheet, DrawingNumber, DrawingPrefix)
		Else
			MessageBox.Show("Either the Drawing Number or Prefix was missing on this sheet, Exiting!")
			Exit Sub
		End If
	Next
MessageBox.Show("Done!")
End Sub

Sub PrintSheet(ByVal sht As Sheet, ByVal DrawingNum As String, ByVal DrawingPrefix As String)
'	Try
	Logger.Debug("begin pdf creation for: " & DrawingNum & "-" & DrawingPrefix)
	Dim oPath As String = ThisDoc.Path
	'Dim PN As String = iProperties.Value("Project", "Part Number")

	'path_and_namePDF = ThisDoc.Pathandname(False)
	Dim oFileName As String = ThisDoc.FileName(False) 'without extension
	Dim oPDFAddIn As ApplicationAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
	Dim oDocument As Document = ThisApplication.ActiveDocument
	Dim oContext As TranslationContext = ThisApplication.TransientObjects.CreateTranslationContext
	oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
	Dim oOptions As NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap
	Dim oDataMedium As DataMedium = ThisApplication.TransientObjects.CreateDataMedium
	If oPDFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
		oOptions.Value("All_Color_AS_Black") = 1
		oOptions.Value("Remove_Line_Weights") = 1
		oOptions.Value("Vector_Resolution") = 400
		oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintCurrentSheet
	End If

	oDataMedium.FileName = oPath & "\" & DrawingNum & " " & DrawingPrefix & ".pdf"
	logger.Debug("PDF Filename: " & odatamedium.FileName)
	'Confirmation message
	ThisApplication.StatusBarText = "PDF SAVED TO: " & oDataMedium.FileName
	'MessageBox.Show("PDF SAVED TO: " & oDataMedium.FileName, "PDF Saved", MessageBoxButtons.OK)

	On Error GoTo handlePDFLock
	'Publish document.
	Call oPDFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
	'--------------------------------------------------------------------------------------------------------------------
	Exit Sub
	handlePDFLock: 
	MessageBox.Show("PDF could not be saved, most likely you or someone else has it open", "No PDF for you " & ThisApplication.GeneralOptions.UserName & "!")
	Resume Next

	handleXLSLock: 
	MessageBox.Show("No XLS", "iLogic")
	Resume Next
'	Catch ex As Exception
'		MessageBox.Show(ex.Message)
'	End Try
End Sub
