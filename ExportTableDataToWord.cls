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

'Word Constant ContentControl name values:
Const cBasicPartsCCName As String = "BasicParts"
Const cAdvancedPartsCCName As String = "AdvancedParts"
Const cAssemblyCCName As String = "AssemblySubjects"
Const cDetailingCCName As String = "Detailing"
Const cPresentationsCCName As String = "PresentationSubjects"
Const cDataManagementCCName As String = "DataManagementSubjects"
Const cOtherCCName As String = "OtherSubjects"
Const cInventorModulesCCName As String = "InventorModules"
Const cOtherFeaturesCCName As String = "OtherFeatures"
Const cBestPracticesCCName As String = "BestPractices"
Const cWhatsNewCCName As String = "WhatsNew"

'Excel rows - we populate these later:
Public sketchingRow As Integer
Public PartRow As Integer
Public AssemblyRow As Integer
Public DetailingRow As Integer
Public PresentationRow As Integer
Public DataManagementRow As Integer
Public OtherRow As Integer
Public InventorModulesRow As Integer
Public OthersRow As Integer
Public BestPracticesRow As Integer
Public WhatsNewRow As Integer

Public weCreatedWord As Boolean




Sub Export_Table_Data_Word()
    weCreatedWord = False
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
    
    'Set wdApp = New Word.Application
    Set wdApp = CreateWord(True)
    'for debugging only:
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Add(wbBook.Path & "\" & stWordDocument)
    
    'may need this in case changing the LockContentControl change doesn't work!
    Call EnterExitDesignMode(wdApp, True)
    
    Call GetSectionRows(ws.usedRange, sketchingRow, PartRow, AssemblyRow, DetailingRow, PresentationRow, DataManagementRow, OtherRow, InventorModulesRow, OthersRow, BestPracticesRow, WhatsNewRow)
    'for debugging purposes:
'    MsgBox ("sketching row: " & CStr(sketchingRow) & vbCrLf & "parts row: " & CStr(PartRow) & vbCrLf & _
'            "Assembly row: " & CStr(AssemblyRow) & vbCrLf & "Detailing Row: " & CStr(DetailingRow) & vbCrLf & _
'            "Presentation Row: " & CStr(PresentationRow) & vbCrLf & "Data Management Row: " & CStr(DataManagementRow) & vbCrLf & _
'            "Other Row: " & CStr(OtherRow) & vbCrLf & "Inventor Modules Row: " & CStr(InventorModulesRow) & vbCrLf & _
'            "Others Row: " & CStr(OthersRow) & vbCrLf & "Best practices: " & CStr(BestPracticesRow) & vbCrLf & "What's New: " & CStr(WhatsNewRow))
    
    Dim r As Range
    Set r = ws.usedRange
    Dim yesnocol As Integer
    yesnocol = 13 'column M
    Dim exportRange As Range
    'this should work but doesn't of course!
    'Set exportRange = r.Find(What:="ExportYes", LookIn:=xlValues, LookAt:=XlLookAt.xlWhole, SearchOrder:=XlSearchOrder.xlByColumns, MatchCase:=False, MatchByte:=False)
    Dim Count As Integer
    Count = 1
    
    Dim i As Integer
    Dim wordTableDataRowStart As Integer
    wordTableDataRowStart = 2
    Dim tmpCControl As ContentControl
    Dim lastTableHeader As String
    For i = 1 To r.Rows.Count
        If Not r.Cells(i, 2).Value = "" Then
            Dim tmpRange As Range
            Set tmpRange = r.Range(r.Cells(i, 1), r.Cells(i, r.Columns.Count))
            Dim foundrange As Range
            Set foundrange = tmpRange.Find(What:="ExportYes", LookIn:=xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=False, MatchByte:=False)
            If Not foundrange Is Nothing Then
                Dim wdTable As Word.Table
                Dim topicCell As Word.Cell
                Dim descrCell As Word.Cell
                Set wdTable = tmpTable(wdDoc, i, r.Rows.Count)
                If lastTableHeader = "" Then
                    lastTableHeader = wdTable.Cell(1, 1).Range.Text
                ElseIf Not lastTableHeader = wdTable.Cell(1, 1).Range.Text Then
                    wordTableDataRowStart = 2
                    lastTableHeader = wdTable.Cell(1, 1).Range.Text
                Else
                    wordTableDataRowStart = wordTableDataRowStart + 1
                    wdTable.Rows.Add (wdTable.Rows(wdTable.Rows.Count - 1))
                End If
                'Set vaData = wdApp.ActiveDocument.Tables(1).Range.Cells(count)
                Set topicCell = wdTable.Cell(wordTableDataRowStart, 1)
                Set descrCell = wdTable.Cell(wordTableDataRowStart, 2)
                
                Dim topicStr As String
                topicStr = ws.Cells(i, 1).Value
                Dim descrStr As String
                descrStr = ws.Cells(i, 2).Value
                
                topicCell.Range.Text = topicStr  'insert value from Excel column A
                descrCell.Range.Text = descrStr 'insert value from Excel column B
                Count = Count + 1
            End If
            'determines which table we need to add the rows to:
        End If
    Next i
    
    lnCountItems = 1
    
    'may need this in case changing the LockContentControl change doesn't work!
    Call EnterExitDesignMode(wdApp, False)
    
    'Save and close the Word doc.
    With wdDoc
        .SaveAs (wbBook.Path & "\Training Survey.docx")
        .Close
    End With
    
    If weCreatedWord Then
        wdApp.Quit
    End If
    
    'Null out the variables.
    Set wdCell = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set r = Nothing
    Set ws = Nothing
    Set wbBook = Nothing
    
    MsgBox "The " & stWordDocument & "'s table has succcessfully " & vbNewLine & _
           "been updated!", vbInformation

End Sub

Public Function CreateWord(Optional bVisible As Boolean = True) As Object

    Dim oTempWD As Object

    On Error Resume Next
    Set oTempWD = GetObject(, "Word.Application")

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ERROR_HANDLER
        Set oTempWD = CreateObject("Word.Application")
        weCreatedWord = True
    End If

    oTempWD.Visible = bVisible
    Set CreateWord = oTempWD

    On Error GoTo 0
    Exit Function

ERROR_HANDLER:
    Select Case Err.Number

        Case Else
            MsgBox "Error " & Err.Number & vbCr & _
                " (" & Err.Description & ") in procedure CreateWord."
            Err.Clear
    End Select

End Function

Function tmpTable(ByVal wdDoc As Word.Document, ByVal excelRow As Integer, ByVal UsedRowCount) As Word.Table
    Dim tmpControl As ContentControl
    Select Case excelRow
        Case sketchingRow To PartRow
            Set tmpControl = findNamedContentControl(wdDoc, cBasicPartsCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case PartRow To AssemblyRow
            Set tmpControl = findNamedContentControl(wdDoc, cAdvancedPartsCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case AssemblyRow To DetailingRow
            Set tmpControl = findNamedContentControl(wdDoc, cAssemblyCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case DetailingRow To PresentationRow
            Set tmpControl = findNamedContentControl(wdDoc, cDetailingCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case PresentationRow To DataManagementRow
            Set tmpControl = findNamedContentControl(wdDoc, cPresentationsCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case DataManagementRow To OtherRow
            Set tmpControl = findNamedContentControl(wdDoc, cDataManagementCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case OtherRow To InventorModulesRow
            Set tmpControl = findNamedContentControl(wdDoc, cOtherCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case InventorModulesRow To OthersRow
            Set tmpControl = findNamedContentControl(wdDoc, cInventorModulesCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case OthersRow To BestPracticesRow
            Set tmpControl = findNamedContentControl(wdDoc, cOtherFeaturesCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case BestPracticesRow To WhatsNewRow
            Set tmpControl = findNamedContentControl(wdDoc, cBestPracticesCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case WhatsNewRow To UsedRowCount
            Set tmpControl = findNamedContentControl(wdDoc, cWhatsNewCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case Else
            'do nothing
    End Select
End Function

'Sub testEnter()
'    EnterExitDesignMode True
'End Sub
'
'Sub testExit()
'    EnterExitDesignMode False
'End Sub

Sub EnterExitDesignMode(ByVal wdApp As Word.Application, bEnter As Boolean)
Dim cbrs As CommandBars
Const sMsoName As String = "DesignMode"

    Set cbrs = wdApp.CommandBars
    If Not cbrs Is Nothing Then
        If cbrs.GetEnabledMso(sMsoName) Then
            If bEnter <> cbrs.GetPressedMso(sMsoName) Then
                cbrs.ExecuteMso sMsoName
'                Stop
            End If
        End If
    End If
End Sub

Sub GetSectionRows(ByVal usedRange As Range, ByRef sketchingRow As Integer, ByRef PartRow As Integer, ByRef AssemblyRow As Integer, ByRef DetailingRow As Integer, ByRef PresentationRow As Integer, _
                        ByRef DataManagementRow As Integer, ByRef OtherRow As Integer, ByRef InventorModulesRow As Integer, ByRef OthersRow As Integer, ByRef BestPracticesRow As Integer, ByRef WhatsNewRow As Integer)

Dim searchRangeStart As Range: Set searchRangeStart = usedRange.Cells(1, 1)
Dim searchRangeEnd As Range: Set searchRangeEnd = usedRange.Cells(usedRange.Rows.Count, 1)

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
    For i = 1 To sectionSearchRange.Rows.Count
        If sectionSearchRange.Cells(i, 1).Value = searchterm Then
            searchForSectionRangeByTerm = sectionSearchRange.Range(i, 1) ' SectionSearchRange.Cells(i, 1).Value
        End If
    Next i
End Function

Function findNamedContentControl(wordDoc As Document, controlName As String) As ContentControl
    Dim Occ As ContentControl
    For Each Occ In wordDoc.ContentControls
        If Occ.Title = controlName Then
            Set findNamedContentControl = Occ
            Exit For
        End If
    Next Occ
End Function

'Function findTableinContentControl(contentCont As ContentControl) As Table
'    findTableinContentControl = contentCont.Range.Tables(1)
'End Function

