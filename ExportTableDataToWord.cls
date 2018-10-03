VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'Constant values = don't change these unless the values in Excel are updated!
Const cSketchingSectionValue As String = "Sketching"
Const cPartModellingSectionValue As String = "Part Modelling"
Const cAssemblyModellingSectionValue As String = "Assembly Modelling"
Const cDetailingSectionValue As String = "Detailing"
Const cPresentationSectionValue As String = "Presentation"
Const cDataManagementSectionValue As String = "Data Management"
Const cOtherSectionValue As String = "Other"
Const cInventorModulesSectionValue As String = "Inventor Modules"
Const cOtherFeaturesSectionValue As String = "Other Features"
Const cBestPracticesSectionValue As String = "Best Practices"
Const cWhatsNewSectionValue As String = "Whats New"


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
    
    'Excel rows:
    Dim sketchingRow As Integer
    Dim PartRow As Integer
    Dim AssemblyRow As Integer
    Dim DetailingRow As Integer
    Dim PresentationRow As Integer
    Dim DataManagementRow As Integer
    Dim OtherRow As Integer
    Dim InventorModulesRow As Integer
    Dim OthersRow As Integer
    Dim BestPracticesRow As Integer
    Dim WhatsNewRow As Integer
    
    Call GetSectionRows(ws.usedRange, sketchingRow, PartRow, AssemblyRow, DetailingRow, PresentationRow, DataManagementRow, OtherRow, InventorModulesRow, OthersRow, BestPracticesRow, WhatsNewRow)
    'for debugging purposes:
    MsgBox ("sketching row: " & CStr(sketchingRow) & vbCrLf & "parts row: " & CStr(PartRow) & vbCrLf & _
            "Assembly row: " & CStr(AssemblyRow) & vbCrLf & "Detailing Row: " & CStr(DetailingRow) & vbCrLf & _
            "Presentation Row: " & CStr(PresentationRow) & vbCrLf & "Data Management Row: " & CStr(DataManagementRow) & vbCrLf & _
            "Other Row: " & CStr(OtherRow) & vbCrLf & "Inventor Modules Row: " & CStr(InventorModulesRow) & vbCrLf & _
            "Others Row: " & CStr(OthersRow) & vbCrLf & "Best practices: " & CStr(BestPracticesRow) & vbCrLf & "What's New: " & CStr(WhatsNewRow))
    
    Dim r As Range
    Set r = ws.usedRange
    Dim yesnocol As Integer
    yesnocol = 13 'column M
    Dim exportRange As Range
    'this should work but doesn't of course!
    'Set exportRange = r.Find(What:="ExportYes", LookIn:=xlValues, LookAt:=XlLookAt.xlWhole, SearchOrder:=XlSearchOrder.xlByColumns, MatchCase:=False, MatchByte:=False)
    Dim count As Integer
    count = 1
    ' this will need amending to either case > number or case < number or if statements
    Dim i As Integer
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
                'Set vaData = wdApp.ActiveDocument.Tables(1).Range.Cells(count)
                Set vaData = wdApp.ActiveDocument.Tables(1).Cell(2, 1)
                Dim descrCell As String
                descrCell = ws.Cells(i, 2).Value
                vaData.Range.InsertAfter descrCell 'insert value from Excel column B
                count = count + 1
            End If
            'determines which table we need to add the rows to:
            Select Case i
                Case sketchingRow To PartRow
                
                Case PartRow To AssemblyRow
                
                Case AssemblyRow To DetailingRow
                
                Case DetailingRow To PresentationRow
                
                Case PresentationRow To DataManagementRow
                
                Case DataManagementRow To OtherRow
                
                Case OtherRow To InventorModulesRow
                
                Case InventorModulesRow To OthersRow
                
                Case OthersRow To BestPracticesRow
                
                Case BestPracticesRow To WhatsNewRow
                
                Case WhatsNewRow To ws.usedRange.Rows.count
                
                Case Else
                    'do nothing
            End Select
            
            
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


Sub GetSectionRows(ByVal usedRange As Range, ByRef sketchingRow As Integer, ByRef PartRow As Integer, ByRef AssemblyRow As Integer, ByRef DetailingRow As Integer, ByRef PresentationRow As Integer, _
                        ByRef DataManagementRow As Integer, ByRef OtherRow As Integer, ByRef InventorModulesRow As Integer, ByRef OthersRow As Integer, ByRef BestPracticesRow As Integer, ByRef WhatsNewRow As Integer)

Dim searchRangeStart As Range: Set searchRangeStart = usedRange.Cells(1, 1)
Dim searchRangeEnd As Range: Set searchRangeEnd = usedRange.Cells(usedRange.Rows.count, 1)

Dim sectionSearchRange As Range: Set sectionSearchRange = usedRange.Range(searchRangeStart, searchRangeEnd)

'section rows:
sketchingRow = searchForSectionRange(sectionSearchRange, cSketchingSectionValue)
PartRow = searchForSectionRange(sectionSearchRange, cPartModellingSectionValue)
AssemblyRow = searchForSectionRange(sectionSearchRange, cAssemblyModellingSectionValue)
DetailingRow = searchForSectionRange(sectionSearchRange, cDetailingSectionValue)
PresentationRow = searchForSectionRange(sectionSearchRange, cPresentationSectionValue)
DataManagementRow = searchForSectionRange(sectionSearchRange, cDataManagementSectionValue)
OtherRow = searchForSectionRange(sectionSearchRange, cOtherSectionValue)
InventorModulesRow = searchForSectionRange(sectionSearchRange, cInventorModulesSectionValue)
OthersRow = searchForSectionRange(sectionSearchRange, cOtherFeaturesSectionValue)
BestPracticesRow = searchForSectionRange(sectionSearchRange, cBestPracticesSectionValue)
WhatsNewRow = searchForSectionRange(sectionSearchRange, cWhatsNewSectionValue)


End Sub

Function searchForSectionRange(ByVal searchrange As Range, searchterm As String) As Integer
    Dim nextRowRange As Range: Set nextRowRange = searchrange.Find(searchterm, LookIn:=xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=True, MatchByte:=False)
    If nextRowRange Is Nothing Then
        searchForSectionRange = searchForSectionRangeByTerm(searchrange, cSketchingSectionValue).row
    Else
        Dim row As Integer: row = 0
        searchForSectionRange = nextRowRange.row
    End If
End Function

Function searchForSectionRangeByTerm(sectionSearchRange As Range, searchterm As String) As Range
    For i = 1 To sectionSearchRange.Rows.count
        If sectionSearchRange.Cells(i, 1).Value = searchterm Then
            searchForSectionRangeByTerm = sectionSearchRange.Range(i, 1) ' SectionSearchRange.Cells(i, 1).Value
        End If
    Next i
End Function


