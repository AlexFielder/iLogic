Attribute VB_Name = "ClientGraphicsText"
Public Sub ClientGraphicsText()
    ' Set a reference to the document.  This will work with
    ' either a part or assembly document.
    Dim oDoc As Document
    Set oDoc = ThisApplication.ActiveDocument

    ' Set a reference to the component definition.
    Dim oCompDef As ComponentDefinition
    Set oCompDef = oDoc.ComponentDefinition

    ' Attempt to get the existing client graphics object.  If it exists
    ' delete it so the rest of the code can continue as if it never existed.
    Dim oClientGraphics As ClientGraphics
    On Error Resume Next
    Set oClientGraphics = oCompDef.ClientGraphicsCollection.Item("Text Test")
    If Err.Number = 0 Then
        oClientGraphics.Delete
    End If
    On Error GoTo 0
    ThisApplication.ActiveView.Update

    ' Create a new ClientGraphics object.
    Set oClientGraphics = oCompDef.ClientGraphicsCollection.Add("Text Test")

    ' Create a graphics node.
    Dim oNode As GraphicsNode
    Set oNode = oClientGraphics.AddNode(1)

    ' Create text graphics.
    Dim oTextGraphics As TextGraphics
    Set oTextGraphics = oNode.AddTextGraphics

    ' Set the properties of the text.
    oTextGraphics.Text = "This is the sample text."
    oTextGraphics.Anchor = ThisApplication.TransientGeometry.CreatePoint(0, 0, 0)
    oTextGraphics.Bold = True
    oTextGraphics.Font = "Arial"
    oTextGraphics.FontSize = 40
    oTextGraphics.HorizontalAlignment = kAlignTextLeft
    oTextGraphics.Italic = True
    Call oTextGraphics.PutTextColor(0, 255, 0)
    oTextGraphics.VerticalAlignment = kAlignTextMiddle

    ' Update the view to see the text.
    ThisApplication.ActiveView.Update
End Sub

