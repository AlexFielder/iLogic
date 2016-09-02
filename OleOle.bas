Attribute VB_Name = "OleOle"
Public Sub addolereferences()
Dim doc As Document
Set doc = ThisApplication.ActiveDocument

'Verify the current document has been saved.
If doc.FullFileName = "" Then
    MsgBox ("This document must be saved first.")
    Exit Sub
End If

Dim selectedfile As String
'Set selectedfile = ""
Dim oFileDlg As Inventor.FileDialog
Set oFileDlg = Nothing
Call ThisApplication.CreateFileDialog(oFileDlg)
oFileDlg.filter = "Dwg files (*.dwg)|*.dwg|Excel files (*.xlsx)|*.xlsx|pdf files (*.pdf)|*.pdf|Inventor parts (*.ipt)|*.ipt|Inventor iFeatures (*.ide)|*.ide"
oFileDlg.InitialDirectory = oOrigRefName
oFileDlg.CancelError = True
oFileDlg.MultiSelectEnabled = True
'On Error Resume Next
Call oFileDlg.ShowOpen
If Err.Number <> 0 Then
Return
ElseIf oFileDlg.FileName <> "" Then
selectedfile = oFileDlg.FileName
End If

'MessageBox.Show("You selected: " & selectedfile , "iLogic")
Dim oleReference As ReferencedOLEFileDescriptor
If InStr(1, selectedfile, "|") > 0 Then
    'If selectedfile.Contains("|") Then ' we have multiple files selected.
    Dim file() As String
    file() = Split(selectedfile, "|")
        For Each s In file
            'MessageBox.Show("You selected: " & s , "iLogic")
            Set oleReference = doc.ReferencedOLEFileDescriptors.Add(s, kOLEDocumentLinkObject)
            oleReference.BrowserVisible = True
            oleReference.Visible = False
            oleReference.DisplayName = Mid$(s, InStrRev(s, "\") + 1)
        Next
    Else
        Set oleReference = doc.ReferencedOLEFileDescriptors.Add(selectedfile, kOLEDocumentLinkObject)
        oleReference.BrowserVisible = True
        oleReference.Visible = False
        oleReference.DisplayName = Mid$(selectedfile, InStrRev(selectedfile, "\") + 1)
    End If
End Sub

Public Sub InsertConstraint()
    ' Set a reference to the assembly component definintion.
    Dim oAsmCompDef As AssemblyComponentDefinition
    Set oAsmCompDef = ThisApplication.ActiveDocument.ComponentDefinition
    
    MsgBox ("Constraint count is: " & oAsmCompDef.Constraints.Count)
'    ' Set a reference to the select set.
'    Dim oSelectSet As SelectSet
'    Set oSelectSet = ThisApplication.ActiveDocument.SelectSet
'
'    ' Validate the correct data is in the select set.
'    If oSelectSet.Count <> 2 Then
'        MsgBox "You must select the two circular edges for the insert."
'        Exit Sub
'    End If
'
'    If Not TypeOf oSelectSet.Item(1) Is Edge Or Not TypeOf oSelectSet.Item(2) Is Edge Then
'        MsgBox "You must select the two circular edges for the insert."
'        Exit Sub
'    End If
'
'    ' Get the two edges from the select set.
'    Dim oEdge1 As Edge
'    Dim oEdge2 As Edge
'    Set oEdge1 = oSelectSet.Item(1)
'    Set oEdge2 = oSelectSet.Item(2)
'
'    ' Create the insert constraint between the parts.
'    Dim oInsert As InsertConstraint
'    Set oInsert = oAsmCompDef.Constraints.AddInsertConstraint(oEdge1, oEdge2, True, 0)
End Sub

