Attribute VB_Name = "WordHelpers"
Option Explicit

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

Function IsWordDocOpen(ByVal wordApp As Word.Application, ByVal Filename As String) As Boolean
    Dim TargetWordDoc As Word.Document

    Dim IteratorWorddoc As Word.Document
    For Each IteratorWorddoc In wordApp.Documents
        If IteratorWorddoc.FullName = Filename Then
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

Sub CloseWord(ByVal wdapp As Word.Application, ByVal wddoc As Word.Document, ByVal wbBook As Workbook, ByVal FileNameToSave As String)
    'Save and close the Word doc.
    With wddoc
        Dim readOnlyDoc As Boolean
        readOnlyDoc = IsWordDocOpen(wdapp, wbBook.path & "\" & FileNameToSave)
        If Not readOnlyDoc Then
            .SaveAs (wbBook.path & "\" & FileNameToSave)
        Else
            MsgBox "Unable to save the resultant document" & vbCrLf & _
            "Perhaps you or someone else has it open?", vbCritical, "Office 3shit5 strikes again!"
            wdapp.Visible = True
        End If
        If Not weAreDebugging And Not readOnlyDoc Then
            .Close
        End If
    End With
    
    If weCreatedWord And Not weAreDebugging And Not readOnlyDoc Then
        wdapp.Quit
    End If
End Sub

'content control-specific:

Sub EnterExitDesignMode(ByVal wdapp As Word.Application, bEnter As Boolean)
Dim cbrs As CommandBars
Const sMsoName As String = "DesignMode"

    Set cbrs = wdapp.CommandBars
    If Not cbrs Is Nothing Then
        If cbrs.GetEnabledMso(sMsoName) Then
            If bEnter <> cbrs.GetPressedMso(sMsoName) Then
                cbrs.ExecuteMso sMsoName
'                Stop
            End If
        End If
    End If
End Sub

Function findNamedContentControl(wordDoc As Document, controlName As String) As ContentControl
    Dim Occ As ContentControl
    For Each Occ In wordDoc.ContentControls
        If Occ.Title = controlName Then
            Set findNamedContentControl = Occ
            Exit For
        End If
    Next Occ
End Function

Public Sub DeleteUnusedCControlByName(ByVal cControlsection As ContentControl, ByVal DeleteSection As Boolean)
    If Not DeleteSection = True Then
        cControlsection.Delete
    End If
End Sub
