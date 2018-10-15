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
Const cPartModellingSectionValue As String = "Part Modelling"
Const cAssemblyModellingSectionValue As String = "Assembly Modelling"
Const cDetailingSectionValue As String = "Detailing"
Const cDataManagementSectionValue As String = "Data Management"
Const cOtherSectionValue As String = "iLogic"
Const cInventorModulesSectionValue As String = "Inventor Modules"
Const cOtherFeaturesSectionValue As String = "Other Features"
Const cWhatsNewSectionValue As String = "Whats New"

'Word Constant ContentControl name values:
Const cBasicPartsCCName As String = "PartSubjects"
Const cAssemblyCCName As String = "AssemblySubjects"
Const cDetailingCCName As String = "Detailing"
Const cDataManagementCCName As String = "DataManagementSubjects"
Const ciLogicrCCName As String = "iLogicSubjects"
Const cInventorModulesCCName As String = "InventorModules"
Const cOtherFeaturesCCName As String = "OtherFeatures"
Const cWhatsNewCCName As String = "WhatsNew"

'Excel rows - we populate these later:
Public PartRow As Integer
Public AssemblyRow As Integer
Public DetailingRow As Integer
Public DataManagementRow As Integer
Public iLogicRow As Integer
Public InventorModulesRow As Integer
Public OthersRow As Integer
Public WhatsNewRow As Integer

Public weCreatedWord As Boolean
Public weAreDebugging As Boolean



Sub Export_Table_Data_Word()
    weCreatedWord = False
    'change this when we're done with development
    weAreDebugging = True
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
    
    Call GetSectionRows(ws.usedRange, PartRow, AssemblyRow, DetailingRow, DataManagementRow, iLogicRow, InventorModulesRow, OthersRow, WhatsNewRow)
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
                Dim TotalDaysCell As Word.Cell
                Dim sectionDataRowStart As Integer
                Dim sectionDataRowEnd As Integer
                Dim SumRange As Range
                
                Set wdTable = tmpTable(wdDoc, i, r.Rows.Count, sectionDataRowStart, sectionDataRowEnd)
                Set SumRange = ws.Cells(sectionDataRowStart, 9) ' column I
                SumRange.Formula = "=SUM(" & Range(Cells(sectionDataRowStart + 2, 9), Cells(sectionDataRowEnd - 1, 9)).Address(False, False) & ")/60/6.5" '.Value = Excel.WorksheetFunction.Sum(ws.Range(sectionDataRowStart, sectionDataRowEnd))
                Debug.Print "SumRange.Value: " & SumRange.Value
                If lastTableHeader = "" Then
                    lastTableHeader = wdTable.Cell(1, 1).Range.Text
                ElseIf Not lastTableHeader = wdTable.Cell(1, 1).Range.Text Then
                    wordTableDataRowStart = 2
                    lastTableHeader = wdTable.Cell(1, 1).Range.Text
                Else
                    wordTableDataRowStart = wordTableDataRowStart + 1
                    wdTable.Rows.Add (wdTable.Rows(wdTable.Rows.Count)) ' - 1))
                End If
                'Set vaData = wdApp.ActiveDocument.Tables(1).Range.Cells(count)
                Set topicCell = wdTable.Cell(wordTableDataRowStart, 1)
                Set descrCell = wdTable.Cell(wordTableDataRowStart, 2)
                Set TotalDaysCell = wdTable.Cell(wdTable.Rows.Count, 2)
                
                Dim topicStr As String
                topicStr = ws.Cells(i, 1).Value
                Dim descrStr As String
                descrStr = ws.Cells(i, 2).Value
                Dim TotalDaysDbl As Double
                TotalDaysDbl = WorksheetFunction.RoundUp((SumRange.Value * 0.02) / 0.5, 0.25) '* 0.5
                Debug.Print "Totaldays rounded using .Text function: " & WorksheetFunction.Text(SumRange.Value, "# ?/?")
                Debug.Print "Totaldays rounded: " & CStr(TotalDaysDbl)
                'returns a message for EVERY row
                'MsgBox ("Total Days for " + lastTableHeader + " : " + CStr(TotalDaysDbl))
                Dim TotalDaysStr As String
                TotalDaysStr = CStr(TotalDaysDbl) ' Round(SumRange.Value2, 1) 'ws.Cells.Formula
                
                If Not topicCell.Shading.BackgroundPatternColor = wdColorAutomatic Then
                    With topicCell
                        .Shading.BackgroundPatternColor = wdColorAutomatic
                        With .Range
                            With .Font
                                .TextColor = wdColorAutomatic
                                .Bold = False
                            End With
                        End With
                    End With
                End If
                
                If Not descrCell.Shading.BackgroundPatternColor = wdColorAutomatic Then
                    With descrCell
                        .Shading.BackgroundPatternColor = wdColorAutomatic
                        With .Range
                            With .Font
                                .TextColor = wdColorAutomatic
                                .Bold = False
                            End With
                        End With
                    End With
                End If
                
                topicCell.Range.Text = topicStr  'insert value from Excel column A
                descrCell.Range.Text = descrStr 'insert value from Excel column B
                ' this will keep getting set and reset so maybe need to only do it we're next to the sectionDataRowEnd value?
                TotalDaysCell.Range.Text = TotalDaysStr
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
        If Not weAreDebugging Then
            .Close
        End If
    End With
    
    If weCreatedWord And Not weAreDebugging Then
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

Function Trunc(vTime As Date) As Date
Dim iHr, iMin As Integer
Dim iQtr    As Integer
Dim MyTime  As Date
    iHr = Hour(vTime)
    iMin = Minute(vTime)
    iQtr = Int(iMin / 15)
    If iMin - (iQtr * 15) < 4 Then
        MyTime = DateAdd("h", iHr, 0)
        MyTime = DateAdd("n", iQtr * 15, MyTime)
    Else
        MyTime = DateAdd("h", iHr, 0)
        MyTime = DateAdd("n", (iQtr + 1) * 15, MyTime)
    End If
    Trunc = MyTime
End Function

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

Function tmpTable(ByVal wdDoc As Word.Document, _
                    ByVal excelRow As Integer, _
                    ByVal UsedRowCount As Integer, _
                    ByRef sectionDataRowStart As Integer, _
                    ByRef sectionDataRowEnd As Integer) As Word.Table
    Dim tmpControl As ContentControl
    Select Case excelRow
        Case PartRow To AssemblyRow
            sectionDataRowStart = PartRow
            sectionDataRowEnd = AssemblyRow
            Set tmpControl = findNamedContentControl(wdDoc, cBasicPartsCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case AssemblyRow To DetailingRow
            sectionDataRowStart = AssemblyRow
            sectionDataRowEnd = DetailingRow
            Set tmpControl = findNamedContentControl(wdDoc, cAssemblyCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case DetailingRow To DataManagementRow
            sectionDataRowStart = DetailingRow
            sectionDataRowEnd = DataManagementRow
            Set tmpControl = findNamedContentControl(wdDoc, cDetailingCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case DataManagementRow To iLogicRow
            sectionDataRowStart = DataManagementRow
            sectionDataRowEnd = iLogicRow
            Set tmpControl = findNamedContentControl(wdDoc, cDataManagementCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case iLogicRow To InventorModulesRow
            sectionDataRowStart = iLogicRow
            sectionDataRowEnd = InventorModulesRow
            Set tmpControl = findNamedContentControl(wdDoc, ciLogicrCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case InventorModulesRow To OthersRow
            sectionDataRowStart = InventorModulesRow
            sectionDataRowEnd = OthersRow
            Set tmpControl = findNamedContentControl(wdDoc, cInventorModulesCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case OthersRow To WhatsNewRow
            sectionDataRowStart = OthersRow
            sectionDataRowEnd = WhatsNewRow
            Set tmpControl = findNamedContentControl(wdDoc, cOtherFeaturesCCName)
            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case WhatsNewRow To UsedRowCount
            sectionDataRowStart = WhatsNewRow
            sectionDataRowEnd = UsedRowCount
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

Sub GetSectionRows(ByVal usedRange As Range, ByRef PartRow As Integer, ByRef AssemblyRow As Integer, ByRef DetailingRow As Integer, ByRef DataManagementRow As Integer, _
                    ByRef iLogicRow As Integer, ByRef InventorModulesRow As Integer, ByRef OthersRow As Integer, ByRef WhatsNewRow As Integer)

Dim searchRangeStart As Range: Set searchRangeStart = usedRange.Cells(1, 1)
Dim searchRangeEnd As Range: Set searchRangeEnd = usedRange.Cells(usedRange.Rows.Count, 1)

Dim sectionSearchRange As Range: Set sectionSearchRange = usedRange.Range(searchRangeStart, searchRangeEnd)

'section rows:
PartRow = searchForSectionRange(sectionSearchRange, cPartModellingSectionValue)
AssemblyRow = searchForSectionRange(sectionSearchRange, cAssemblyModellingSectionValue)
DetailingRow = searchForSectionRange(sectionSearchRange, cDetailingSectionValue)
DataManagementRow = searchForSectionRange(sectionSearchRange, cDataManagementSectionValue)
iLogicRow = searchForSectionRange(sectionSearchRange, cOtherSectionValue)
InventorModulesRow = searchForSectionRange(sectionSearchRange, cInventorModulesSectionValue)
OthersRow = searchForSectionRange(sectionSearchRange, cOtherFeaturesSectionValue)
WhatsNewRow = searchForSectionRange(sectionSearchRange, cWhatsNewSectionValue)


End Sub

Function searchForSectionRange(ByVal searchrange As Range, searchterm As String) As Integer
    Dim nextRowRange As Range: Set nextRowRange = searchrange.Find(searchterm, LookIn:=xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=True, MatchByte:=False)
    If nextRowRange Is Nothing Then
        searchForSectionRange = searchForSectionRangeByTerm(searchrange, cPartModellingSectionValue).row
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

