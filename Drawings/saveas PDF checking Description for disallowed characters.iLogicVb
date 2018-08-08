Option Explicit on
Sub Main()
	
	'Get document file name
'	docname = ThisDoc.FileName(False)
	'Get the model which the drawing is referencing (1st view)
	Dim oModelDoc As String = IO.Path.GetFileName(ThisDrawing.ModelDocument.FullFileName)
	Dim oDesc As String = iProperties.Value(oModelDoc, "Project", "Description")
	Dim disallowedlist As String = "<>:;/\|?*" & Chr(34)
	If oDesc = "" Then
		MessageBox.Show("The Description iproperty cannot be empty. Go to the Part/Assembly and type a Description. NO PDF WILL BE GENERATED", "Warning")
		Exit Sub
	Else
		'debug:
'		MessageBox.Show("Description not empty, continuing!" & (oDesc.IndexOfAny(disallowedlist.ToCharArray) > 1))
		If oDesc.IndexOfAny(disallowedlist.ToCharArray) > -1 Then
			MessageBox.Show("The Description contains an illegal character, exiting!")
			Exit Sub
		End If
	End If

    ' Get the PDF translator Add-In.
    Dim PDFAddIn As TranslatorAddIn
    PDFAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")

    'Set a reference to the active document (the document to be published).
    Dim oDocument As Document
    oDocument = ThisApplication.ActiveDocument

    Dim oContext As TranslationContext
    oContext = ThisApplication.TransientObjects.CreateTranslationContext
    oContext.Type = kFileBrowseIOMechanism

    ' Create a NameValueMap object
    Dim oOptions As NameValueMap
    oOptions = ThisApplication.TransientObjects.CreateNameValueMap

    ' Create a DataMedium object
    Dim oDataMedium As DataMedium
    oDataMedium = ThisApplication.TransientObjects.CreateDataMedium

    ' Check whether the translator has 'SaveCopyAs' options
    If PDFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
        ' Options for drawings...
        oOptions.Value("All_Color_AS_Black") = 0
        oOptions.Value("Remove_Line_Weights") = 0
        oOptions.Value("Vector_Resolution") = 400
        oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
        'oOptions.Value("Custom_Begin_Sheet") = 2
        'oOptions.Value("Custom_End_Sheet") = 4
    End If
    
    'get PDF target folder path
    Dim oPath As String = IO.Path.GetDirectoryName(ThisDrawing.ModelDocument.FullFileName) ' "C:\Users\user\Desktop\test"
	MessageBox.Show(oPath)
    'oFolder = Left(oPath, InStrRev(oPath, "\")) & "PDF"
    Dim oFileName As String = ThisDoc.FileName(False) 'without extension
    Dim oRevNum As String = iProperties.Value("Project", "Revision Number")
	Dim oDate As String = DateTime.Now.ToString("_yyMMdd")
    
	If oDesc = "" Then
		MessageBox.Show("Warning", "Description iproperties is empty")
	End If
	Dim oPathCom As String = String.Empty
	If oRevNum = "-" Or oRevNum = "" Then
		oPathCom = oPath & "\" & oFileName & " - " & oDesc & oDate &".pdf"
	Else
		oPathCom = oPath & "\" & oFileName & "-"  & oRevNum & " - " & oDesc & oDate & ".pdf"
	End If

	oDataMedium.FileName = oPathCom
	

    'Publish document.
    Call PDFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
    '
End Sub