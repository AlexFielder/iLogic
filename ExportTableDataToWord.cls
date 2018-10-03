VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Export_Table_Data_Word()

    'Name of the existing Word document
    Const stWordDocument As String = "Template.dotx" 'this needs to exist in the folder this Excel file is in!
    
    'Word objects.
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    Dim wdCell As Word.Cell
    
    'Excel objects
    Dim wbBook As Workbook
    Dim ws As Worksheet
    
    'Count used in a FOR loop to fill the Word table.
    Dim lnCountItems As Long
    
    'Variant to hold the data to be exported.
    Dim vaData As Variant
    
    'Initialize the Excel objects
    Set wbBook = ThisWorkbook
    Set ws = wbBook.Worksheets("Inventor Survey Beta")
    
    Set wdApp = New Word.Application
    'for debugging only:
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Add(wbBook.Path & "\" & stWordDocument)
    
    Dim r As Range
    Set r = ws.UsedRange
    Dim yesnocol As Integer
    yesnocol = 13 'column M
    Dim exportRange As Range
    'this should work but doesn't of course!
    'Set exportRange = r.Find(What:="ExportYes", LookIn:=xlValues, LookAt:=XlLookAt.xlWhole, SearchOrder:=XlSearchOrder.xlByColumns, MatchCase:=False, MatchByte:=False)
    Dim count As Integer
    count = 1
    For i = 1 To r.Rows.count
        If Not r.Cells(i, 2).Value = "" Then
            Dim tmpRange As Range
            Set tmpRange = r.Range(r.Cells(i, 1), r.Cells(i, r.Columns.count))
            Dim foundrange As Range
            Set foundrange = tmpRange.Find(What:="ExportYes", LookIn:=xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=False, MatchByte:=False)
            If Not foundrange Is Nothing Then
                Dim wdTable As Table
                Set wdTable = wdApp.ActiveDocument.Tables(1)
                wdTable.Rows.Add
                Set vaData = wdApp.ActiveDocument.Tables(1).Range.Cells(count)
                Dim descrCell As String
                descrCell = ws.Cells(i, 2).Value
                vaData.Range.InsertAfter descrCell 'insert value from Excel column B
                count = count + 1
            End If
        End If
    Next i
    
'    vaData = wsSheet.Range("A1:A10").Value
    
    'Instantiate Word and open the "Table Reports" document.
'    Set wdApp = New Word.Application
'    Set wdDoc = wdApp.Documents.Open(wbBook.Path & "\" & stWordDocument)
    
    lnCountItems = 1
    
'    'Place the data from the variant into the table in the Word doc.
'    For Each wdCell In wdDoc.Tables(1).Columns(1).Cells
'        wdCell.Range.Text = vaData(lnCountItems, 1)
'        lnCountItems = lnCountItems + 1
'    Next wdCell
    
    'Save and close the Word doc.
    With wdDoc
        .SaveAs (wbBook.Path & "\Training Survey.docx")
        .Close
    End With
    
    wdApp.Quit
    
    'Null out the variables.
    Set wdCell = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    MsgBox "The " & stWordDocument & "'s table has succcessfully " & vbNewLine & _
           "been updated!", vbInformation

End Sub

