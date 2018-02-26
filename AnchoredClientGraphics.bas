Attribute VB_Name = "AnchoredClientGraphics"
Public Sub AnchoredClientGraphics()
  ' Set a reference to the document.
  Dim oDoc As Document
  Set oDoc = ThisApplication.ActiveDocument

  ' Set a reference to the component definition.
  ' This assumes that the active document is a part or an assembly.
  Dim oCompDef As ComponentDefinition
  Set oCompDef = oDoc.ComponentDefinition

  ' Attempt to get the existing client graphics object. If it exists
  ' delete it so the rest of the code can continue as if it never existed.
  Dim oClientGraphics As ClientGraphics
  On Error Resume Next
  Set oClientGraphics = oCompDef.ClientGraphicsCollection.Item("Anchored Text")
  If Err.Number = 0 Then
    oClientGraphics.Delete
  End If
  On Error GoTo 0
  ThisApplication.ActiveView.Update

  ' Create a new ClientGraphics object.
  Set oClientGraphics = oCompDef.ClientGraphicsCollection.Add("Anchored Text")

  ' Create a graphics node.
  Dim oNode As GraphicsNode
  Set oNode = oClientGraphics.AddNode(1)

  ' Create text graphics.
  Dim oTextGraphics As TextGraphics
  Set oTextGraphics = oNode.AddTextGraphics

  ' Set the properties of the text.
  oTextGraphics.Text = "Anchored text."
  oTextGraphics.Bold = True
  oTextGraphics.FontSize = 30
  Call oTextGraphics.PutTextColor(47, 47, 48) '(117, 76, 45)

  Dim oAnchorPoint As Point
  Set oAnchorPoint = ThisApplication.TransientGeometry.CreatePoint(1, 1, 1)

  ' Set the text's anchor in model space.
  oTextGraphics.Anchor = oAnchorPoint

  ' Anchor the text graphics in the view.
  Call oTextGraphics.SetViewSpaceAnchor( _
      oAnchorPoint, ThisApplication.TransientGeometry.CreatePoint2d(150, 30), ViewLayoutEnum.kBottomRightViewCorner)

  ' Update the view to see the text.
  ThisApplication.ActiveView.Update
End Sub

