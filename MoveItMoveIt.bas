Attribute VB_Name = "MoveItMoveIt"

Public Sub MoveOccurrence()
' Set a reference to the assembly component definintion.
Dim oAsmCompDef As AssemblyComponentDefinition
Set oAsmCompDef = ThisApplication.ActiveDocument.ComponentDefinition
Dim XSpacingInc As Double
XSpacingInc = CDbl(InputBox("X Offset", "X OFffset", "0")) 'default is in cm
Dim YSpacingInc As Double
YSpacingInc = CDbl(InputBox("Y Spacing", "Y Spacing", "30")) 'default is in cm
Dim YSpacing As Double
Dim XSpacing As Double
XSpacing = 0
YSpacing = 0
Dim oSelect As New clsSelect

Dim oSelectedEnts As ObjectsEnumerator
Set oSelectedEnts = oSelect.PickPartToMove(kAssemblyOccurrenceFilter)
'Set oSelectedEnts = ThisApplication.CommandManager.Pick(kAssemblyOccurrenceFilter, "Pick something")
' Get an occurrence from the select set.
For i = 1 To oSelectedEnts.Count
    On Error Resume Next
    Dim oOccurrence As ComponentOccurrence
    Set oOccurrence = oSelectedEnts.Item(i)
    If Err Then
      MsgBox "An occurrence must be selected."
      Exit Sub
    End If
    On Error GoTo 0
    
    ' Get the current transformation matrix from the occurrence.
    Dim oTransform As Matrix
    Set oTransform = oOccurrence.Transformation
    
    ' Move the occurrence honoring any existing constraints.
    oTransform.SetTranslation ThisApplication.TransientGeometry.CreateVector(XSpacing + XSpacingInc, YSpacing + YSpacingInc, 0)
    oOccurrence.Transformation = oTransform
    XSpacing = XSpacing + XSpacingInc
    YSpacing = YSpacing + YSpacingInc
    ' Move the occurrence ignoring any constraints.
    ' Anything that causes the assembly to recompute will cause the
    ' occurrence to reposition itself to honor the constraints.
    'oTransform.SetTranslation ThisApplication.TransientGeometry.CreateVector(3, 4, 5)
    'Call oOccurrence.SetTransformWithoutConstraints(oTransform)
Next i
End Sub

