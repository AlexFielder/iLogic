Attribute VB_Name = "ClientGraphicsPrimitives"
Public Sub ClientGraphicsPrimitives()
    Dim oDoc As Document
    Set oDoc = ThisApplication.ActiveDocument

    ' Set a reference to component definition of the active document.
    ' This assumes that a part or assembly document is active.
    Dim oCompDef As ComponentDefinition
    Set oCompDef = ThisApplication.ActiveDocument.ComponentDefinition

    ' Check to see if the test graphics data object already exists.
    ' If it does clean up by removing all associated of the client graphics
    ' from the document. If it doesn't create it.
    On Error Resume Next
    Dim oClientGraphics As ClientGraphics
    Set oClientGraphics = oCompDef.ClientGraphicsCollection.Item("SampleGraphicsID")
    If Err.Number = 0 Then
        On Error GoTo 0
        ' An existing client graphics object was successfully obtained so clean up.
        oClientGraphics.Delete
        
        ' update the display to see the results.
        ThisApplication.ActiveView.Update
    Else
        Err.Clear
        On Error GoTo 0

        ' Set a reference to the transient geometry object for user later.
        Dim oTransGeom As TransientGeometry
        Set oTransGeom = ThisApplication.TransientGeometry

        ' Create the ClientGraphics object.
        Set oClientGraphics = oCompDef.ClientGraphicsCollection.Add("SampleGraphicsID")

        ' Create a new graphics node within the client graphics objects.
        Dim oCurvesNode As GraphicsNode
        Set oCurvesNode = oClientGraphics.AddNode(1)
        
        Dim oCenter As Point
        Set oCenter = oTransGeom.CreatePoint(1, 1, 0)
        
        Dim oNormal As UnitVector
        Set oNormal = oTransGeom.CreateUnitVector(0, 0, 1)
        
        ' Create a transient circle object
        Dim oCircle As Inventor.Circle
        Set oCircle = oTransGeom.CreateCircle(oCenter, oNormal, 1)
        
        ' Create a circle graphics object within the node.
        Dim oCircleGraphics As CurveGraphics
        Set oCircleGraphics = oCurvesNode.AddCurveGraphics(oCircle)
       
        Dim oReference As UnitVector
        Set oReference = oTransGeom.CreateUnitVector(-1, 0, 0)
        
        ' Create a transient arc object
        Dim oArc1 As Arc3d
        Set oArc1 = oTransGeom.CreateArc3d(oCenter, oNormal, oReference, 3, 3.14159 / 4, 3.14159 / 2)
        
        ' Create an arc graphics object within the node.
        Dim oArcGraphics1 As CurveGraphics
        Set oArcGraphics1 = oCurvesNode.AddCurveGraphics(oArc1)
        
        ' Create a transient arc object
        Dim oArc2 As Arc3d
        Set oArc2 = oTransGeom.CreateArc3d(oCenter, oNormal, oReference, 3, 3.14159 + (3.14159 / 4), 3.14159 / 2)
        
        ' Create an arc graphics object within the node.
        Dim oArcGraphics2 As CurveGraphics
        Set oArcGraphics2 = oCurvesNode.AddCurveGraphics(oArc2)
        
        ' Create a transient line segment object
        Dim oLineSegment1 As LineSegment
        Set oLineSegment1 = oTransGeom.CreateLineSegment(oArc1.StartPoint, oArc2.EndPoint)
        
        ' Create an line graphics object within the node.
        Dim oLineGraphics1 As CurveGraphics
        Set oLineGraphics1 = oCurvesNode.AddCurveGraphics(oLineSegment1)
        
        ' Create a transient line segment object
        Dim oLineSegment2 As LineSegment
        Set oLineSegment2 = oTransGeom.CreateLineSegment(oArc2.StartPoint, oArc1.EndPoint)
        
        ' Create an line graphics object within the node.
        Dim oLineGraphics2 As CurveGraphics
        Set oLineGraphics2 = oCurvesNode.AddCurveGraphics(oLineSegment2)
        
        ' Update the view to see the resulting curves.
        ThisApplication.ActiveView.Update
    End If
End Sub

