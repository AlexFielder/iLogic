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
Const cTotalTrainingDaysSectionValue As String = "Total Training Days"

'Word Constant ContentControl name values:
Const cBasicPartsCCName As String = "PartSubjects"
Const cAssemblyCCName As String = "AssemblySubjects"
Const cDetailingCCName As String = "Detailing"
Const cDataManagementCCName As String = "DataManagementSubjects"
Const ciLogicrCCName As String = "iLogicSubjects"
Const cInventorModulesCCName As String = "InventorModules"
Const cOtherFeaturesCCName As String = "OtherFeatures"
Const cWhatsNewCCName As String = "WhatsNew"
Const cTotalTrainingDaysCCName As String = "TotalDays"

'Excel rows - we populate these later:
Public PartRow As Integer
Public AssemblyRow As Integer
Public DetailingRow As Integer
Public DataManagementRow As Integer
Public iLogicRow As Integer
Public InventorModulesRow As Integer
Public OthersRow As Integer
Public WhatsNewRow As Integer
Public TotalTrainingDaysRow As Integer

Public TotalTrainingDays As Double

'Boolean Values
Public weCreatedWord As Boolean
Public weAreDebugging As Boolean
'Booleans for section deletion
Public partsRequired As Boolean
Public assembliesRequired As Boolean
Public detailingRequired As Boolean
Public dataManagementRequired As Boolean
Public iLogicRequired As Boolean
Public inventorModulesRequired As Boolean
Public otherFeaturesRequired As Boolean
Public whatsNewRequired As Boolean
Public totalTrainingDaysRequired As Boolean

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
    If weAreDebugging Then
        wdApp.Visible = True
    End If
    Set wdDoc = wdApp.Documents.Add(wbBook.Path & "\" & stWordDocument)
    
    'may need this in case changing the LockContentControl change doesn't work!
    Call EnterExitDesignMode(wdApp, True)
    
    Call SetSectionDefaultRequirements
    
    Call GetSectionRows(ws.usedRange, PartRow, AssemblyRow, DetailingRow, DataManagementRow, iLogicRow, InventorModulesRow, OthersRow, WhatsNewRow, TotalTrainingDaysRow)
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
    Dim ccList As Collection
    Set ccList = New Collection
    
    '(Step 1) Display your Progress Bar
    ufProgress.LabelProgress.Width = 0
    ufProgress.Show
    Dim percentDone As Single
    
    TotalTrainingDays = 0
    
    For i = 1 To TotalTrainingDaysRow ' r.Rows.Count
        percentDone = i / TotalTrainingDaysRow ' r.Rows.Count
        With ufProgress
            .LabelCaption.Caption = "Processing Row " & i & " of " & TotalTrainingDaysRow 'r.Rows.Count
            .LabelProgress.Width = percentDone * (.FrameProgress.Width)
        End With
        DoEvents
        If Not r.Cells(i, 2).Value = "" Then
            Dim tmpRange As Range
            Set tmpRange = r.Range(r.Cells(i, 1), r.Cells(i, r.Columns.Count))
            Dim foundrange As Range
            Set foundrange = tmpRange.Find(What:="ExportYes", LookIn:=xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=False, MatchByte:=False)
            
            Dim wdTable As Word.Table
            Dim sectionDataRowStart As Integer
            Dim sectionDataRowEnd As Integer
            Dim thisControl As ContentControl
            Set wdTable = tmpTable(wdDoc, i, r.Rows.Count, sectionDataRowStart, sectionDataRowEnd, thisControl)
            If Not thisControl Is Nothing Then
                ccList.Add thisControl
            End If
            If Not foundrange Is Nothing Then
                Dim headerCell As Word.Cell
                Dim topicCell As Word.Cell
                Dim descrCell As Word.Cell
                Dim TotalDaysCell As Word.Cell
                
                Dim SumRange As Range
                Call SetSectionRequirementsByName(thisControl)
                Set SumRange = ws.Cells(sectionDataRowStart, 9) ' column I
                SumRange.Formula = "=SUM(" & Range(Cells(sectionDataRowStart + 2, 9), Cells(sectionDataRowEnd - 1, 9)).Address(False, False) & ")/60/6.5" '.Value = Excel.WorksheetFunction.Sum(ws.Range(sectionDataRowStart, sectionDataRowEnd))
                Debug.Print "SumRange.Value: " & SumRange.Value
                
                Dim TotalDaysDbl As Double
                TotalDaysDbl = CStr(Round(SumRange.Value2, 2)) ' WorksheetFunction.RoundUp((SumRange.Value * 0.02) / 0.5, 0.25) '* 0.5
                Debug.Print "Totaldays rounded using .Text function: " & WorksheetFunction.Text(SumRange.Value, "# ?/?")
                Debug.Print "Total days rounded using Round API function: " & CStr(Round(SumRange.Value2, 2))
                Debug.Print "Totaldays rounded: " & CStr(TotalDaysDbl)
                
                'Debug.Print dTime & vbTab & Format(Trunc(dTime), "hh:mm")
                'returns a message for EVERY row
                'MsgBox ("Total Days for " + lastTableHeader + " : " + CStr(TotalDaysDbl))
                Dim TotalDaysStr As String
                TotalDaysStr = CStr(TotalDaysDbl) ' Round(SumRange.Value2, 2)) 'CStr(SumRange.Value) ' Round(SumRange.Value2, 1) 'ws.Cells.Formula
                
                If lastTableHeader = "" Then
                    lastTableHeader = wdTable.Cell(1, 1).Range.Text
                ElseIf Not lastTableHeader = wdTable.Cell(1, 1).Range.Text Then
                    'this should only fire once per table (hopefully!)
                    TotalTrainingDays = TotalTrainingDays + TotalDaysDbl
                    Debug.Print "Total Training Days Running Total= " & CStr(TotalTrainingDays)
                    wordTableDataRowStart = 2
                    lastTableHeader = wdTable.Cell(1, 1).Range.Text
                Else
                    wordTableDataRowStart = wordTableDataRowStart + 1
                    wdTable.Rows.Add (wdTable.Rows(wdTable.Rows.Count)) ' - 1))
                End If
                'Set vaData = wdApp.ActiveDocument.Tables(1).Range.Cells(count)
                Set headerCell = wdTable.Cell(1, 1)
                Set topicCell = wdTable.Cell(wordTableDataRowStart, 1)
'                Set descrCell = wdTable.Cell(wordTableDataRowStart, 2)
                If i = TotalTrainingDaysRow Then
                    'both of these should work
                    'Set TotalDaysCell = wdTable.Cell(wdTable.Rows.Count, 1)
                    Set descrCell = wdTable.Cell(wdTable.Rows.Count, 1)
                Else
                    Set topicCell = wdTable.Cell(wordTableDataRowStart, 1)
                    Set descrCell = wdTable.Cell(wordTableDataRowStart, 2)
                    Set TotalDaysCell = wdTable.Cell(wdTable.Rows.Count, 2)
                End If
                Dim topicStr As String
                topicStr = ws.Cells(i, 1).Value
                Dim descrStr As String
                descrStr = ws.Cells(i, 2).Value
                
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
                
                descrCell.Range.Text = descrStr 'insert value from Excel column B
                If Not i = TotalTrainingDaysRow Then
                    topicCell.Range.Text = topicStr  'insert value from Excel column A
                    TotalDaysCell.Range.Text = TotalDaysStr
                End If
                Count = Count + 1
            End If
            'determines which table we need to add the rows to:
        End If
        If i = r.Rows.Count Then Unload ufProgress
    Next i
    
    lnCountItems = 1
    
    Call DeleteUnusedSections(ccList)
    'update Table of Contents
    If wdDoc.TablesOfContents.Count = 1 Then wdDoc.TablesOfContents(1).Update
    'may need this in case changing the LockContentControl change doesn't work!
    Call EnterExitDesignMode(wdApp, False)
    
    'Save and close the Word doc.
    With wdDoc
        Dim readOnlyDoc As Boolean
        readOnlyDoc = IsWordDocOpen(wdApp, wbBook.Path & "\Training Survey.docx")
        If Not readOnlyDoc Then
            .SaveAs (wbBook.Path & "\Training Survey.docx")
        Else
            MsgBox "Unable to save the resultant document" & vbCrLf & _
            "Perhaps you or someone else has it open?", vbCritical, "Office 3shit5 strikes again!"
            wdApp.Visible = True
        End If
        If Not weAreDebugging And Not readOnlyDoc Then
            .Close
        End If
    End With
    
    If weCreatedWord And Not weAreDebugging And Not readOnlyDoc Then
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

Function IsWorkbookOpen(ByVal FileName As String) As Boolean
    Dim TargetWorkbook As Workbook

    Dim IteratorWorkbook As Workbook
    For Each IteratorWorkbook In Application.Workbooks
        If IteratorWorkbook.FullName = FileName Then
            Set TargetWorkbook = IteratorWorkbook
        End If
    Next

    If Not TargetWorkbook Is Nothing Then
        If TargetWorkbook.ReadOnly Then
            IsWorkbookOpen = True
            Exit Function
        End If
    End If
End Function

Function IsWordDocOpen(ByVal wordApp As Word.Application, ByVal FileName As String) As Boolean
    Dim TargetWordDoc As Word.Document

    Dim IteratorWorddoc As Word.Document
    For Each IteratorWorddoc In wordApp.Documents
        If IteratorWorddoc.FullName = FileName Then
            Set TargetWordDoc = IteratorWorddoc
            IsWordDocOpen = True
            Exit Function
        End If
    Next

    If Not TargetWordDoc Is Nothing Then
        If TargetWordDoc.ReadOnly Then
            IsWordDocOpen = True
            Exit Function
        End If
    End If
End Function

Public Sub DeleteUnusedSections(ByVal contentControlList As Collection)
    Dim controlnum As Integer
    For controlnum = 1 To contentControlList.Count
        Dim cControl As ContentControl
        Set cControl = contentControlList(controlnum)
        Select Case cControl.Title
            Case cBasicPartsCCName
                Call DeleteUnusedCControlByName(cControl, partsRequired)
            Case cAssemblyCCName
                Call DeleteUnusedCControlByName(cControl, assembliesRequired)
            Case cDetailingCCName
                Call DeleteUnusedCControlByName(cControl, detailingRequired)
            Case cDataManagementCCName
                Call DeleteUnusedCControlByName(cControl, dataManagementRequired)
            Case ciLogicrCCName
                Call DeleteUnusedCControlByName(cControl, iLogicRequired)
            Case cInventorModulesCCName
                Call DeleteUnusedCControlByName(cControl, inventorModulesRequired)
            Case cOtherFeaturesCCName
                Call DeleteUnusedCControlByName(cControl, otherFeaturesRequired)
            Case cWhatsNewCCName
                Call DeleteUnusedCControlByName(cControl, whatsNewRequired)
            Case Else
        End Select
    Next
End Sub

Public Sub DeleteUnusedCControlByName(ByVal cControlsection As ContentControl, ByVal DeleteSection As Boolean)
    If Not DeleteSection = True Then
        cControlsection.Delete
    End If
End Sub

Public Sub SetSectionRequirementsByName(ByVal cControlsection As ContentControl)
    Select Case cControlsection.Title
        Case cBasicPartsCCName
            partsRequired = True
        Case cAssemblyCCName
            assembliesRequired = True
        Case cDetailingCCName
            detailingRequired = True
        Case cDataManagementCCName
            dataManagementRequired = True
        Case ciLogicrCCName
            iLogicRequired = True
        Case cInventorModulesCCName
            inventorModulesRequired = True
        Case cOtherFeaturesCCName
            otherFeaturesRequired = True
        Case cWhatsNewCCName
            whatsNewRequired = True
        Case cTotalTrainingDaysCCName
            totalTrainingDaysRequired = True
        Case Else
    End Select
End Sub

Public Sub SetSectionDefaultRequirements()
    partsRequired = False
    assembliesRequired = False
    detailingRequired = False
    dataManagementRequired = False
    iLogicRequired = False
    inventorModulesRequired = False
    otherFeaturesRequired = False
    whatsNewRequired = False
    totalTrainingDaysRequired = True
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
                    ByRef sectionDataRowEnd As Integer, _
                    ByRef thisControl As ContentControl) As Word.Table
    Dim tmpControl As ContentControl
    Select Case excelRow
        Case PartRow To AssemblyRow
            sectionDataRowStart = PartRow
            sectionDataRowEnd = AssemblyRow
            Set tmpControl = findNamedContentControl(wdDoc, cBasicPartsCCName)
'            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case AssemblyRow To DetailingRow
            sectionDataRowStart = AssemblyRow
            sectionDataRowEnd = DetailingRow
            Set tmpControl = findNamedContentControl(wdDoc, cAssemblyCCName)
'            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case DetailingRow To DataManagementRow
            sectionDataRowStart = DetailingRow
            sectionDataRowEnd = DataManagementRow
            Set tmpControl = findNamedContentControl(wdDoc, cDetailingCCName)
'            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case DataManagementRow To iLogicRow
            sectionDataRowStart = DataManagementRow
            sectionDataRowEnd = iLogicRow
            Set tmpControl = findNamedContentControl(wdDoc, cDataManagementCCName)
'            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case iLogicRow To InventorModulesRow
            sectionDataRowStart = iLogicRow
            sectionDataRowEnd = InventorModulesRow
            Set tmpControl = findNamedContentControl(wdDoc, ciLogicrCCName)
'            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case InventorModulesRow To OthersRow
            sectionDataRowStart = InventorModulesRow
            sectionDataRowEnd = OthersRow
            Set tmpControl = findNamedContentControl(wdDoc, cInventorModulesCCName)
'            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case OthersRow To WhatsNewRow
            sectionDataRowStart = OthersRow
            sectionDataRowEnd = WhatsNewRow
            Set tmpControl = findNamedContentControl(wdDoc, cOtherFeaturesCCName)
'            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case WhatsNewRow To TotalTrainingDaysRow - 2
            sectionDataRowStart = WhatsNewRow
            sectionDataRowEnd = TotalTrainingDaysRow
            Set tmpControl = findNamedContentControl(wdDoc, cWhatsNewCCName)
'            tmpControl.LockContentControl = False
            Set tmpTable = tmpControl.Range.Tables(1)
        Case TotalTrainingDaysRow To UsedRowCount
            'we're probably ready to populate the Total Days.
            sectionDataRowStart = TotalTrainingDaysRow
            sectionDataRowEnd = UsedRowCount
            Set tmpControl = findNamedContentControl(wdDoc, cTotalTrainingDaysCCName)
            Set tmpTable = tmpControl.Range.Tables(1)
        Case Else
            'do nothing
    End Select
    If Not tmpControl Is Nothing Then
       Set thisControl = tmpControl
    End If
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
                    ByRef iLogicRow As Integer, ByRef InventorModulesRow As Integer, ByRef OthersRow As Integer, ByRef WhatsNewRow As Integer, ByRef TotalTrainingDaysRow As Integer)

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
TotalTrainingDaysRow = searchForSectionRange(sectionSearchRange, cTotalTrainingDaysSectionValue)

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



