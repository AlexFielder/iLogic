Attribute VB_Name = "DebugObject"
Public Sub DebugObject()
    Dim oDoc As Document
    Set oDoc = ThisApplication.ActiveDocument
    Dim oObj As Object
    Set oObj = oDoc.SelectSet.Item(1)
    Stop
End Sub
