Attribute VB_Name = "DrawingOverlayGraphics"
Dim oIE As InteractionEvents


Public Sub DrawOverlayGraphics()
    Dim oDoc As Document
    Set oDoc = ThisApplication.ActiveDocument

    Set oIE = ThisApplication.CommandManager.CreateInteractionEvents
    oIE.Start
    
    Dim oIG As InteractionGraphics
    Set oIG = oIE.InteractionGraphics
    
    On Error Resume Next
    Dim oDataSets As GraphicsDataSets
    Set oDataSets = oIG.GraphicsDataSets
        
    ' Set a reference to the transient geometry object for use later.
    Dim oTransGeom As TransientGeometry
    Set oTransGeom = ThisApplication.TransientGeometry

    ' Create a coordinate set.
    Dim oCoordSet As GraphicsCoordinateSet
    Set oCoordSet = oDataSets.CreateCoordinateSet(1)
    
    ' Create an array that contains coordinates that define a set
    ' of outwardly spiraling points.
    Dim oPointCoords(1 To 90) As Double
    Dim i As Long
    Dim dRadius As Double
    dRadius = 1
    Dim dAngle As Double
    For i = 0 To 29
        ' Define the X, Y, and Z components of the point.
        oPointCoords(i * 3 + 1) = dRadius * Cos(dAngle)
        oPointCoords(i * 3 + 2) = dRadius * Sin(dAngle)
        oPointCoords(i * 3 + 3) = i / 2
       
        ' Increment the angle and radius to create the spiral.
        dRadius = dRadius + 0.25
        dAngle = dAngle + (3.14159265358979 / 6)
    Next
   
    ' Assign the points into the coordinate set.
    Call oCoordSet.PutCoordinates(oPointCoords)
   
    ' Create the ClientGraphics object.
    Dim oClientGraphics As ClientGraphics
    Set oClientGraphics = oIG.OverlayClientGraphics
   
    ' Create a new graphics node within the client graphics objects.
    Dim oLineNode As GraphicsNode
    Set oLineNode = oClientGraphics.AddNode(1)
   
    ' Create a LineGraphics object within the node.
    Dim oLineSet As LineGraphics
    Set oLineSet = oLineNode.AddLineGraphics
   
    ' Assign the coordinate set to the line graphics.
    oLineSet.CoordinateSet = oCoordSet
   
    ' Update the view to see the resulting spiral.
    oIG.UpdateOverlayGraphics ThisApplication.ActiveView
   
    ' Create another graphics node for a line strip.
    Dim oLineStripNode As GraphicsNode
    Set oLineStripNode = oClientGraphics.AddNode(2)
   
    ' Create a LineStripGraphics object within the new node.
    Dim oLineStrip As LineStripGraphics
    Set oLineStrip = oLineStripNode.AddLineStripGraphics
   
    ' Assign the same coordinate set to the line strip.
    oLineStrip.CoordinateSet = oCoordSet
   
    ' Create a color set to use in defining a explicit color to the line strip.
    Dim oColorSet As GraphicsColorSet
    Set oColorSet = oDataSets.CreateColorSet(1)
   
    ' Add a single color to the set that is red.
    Call oColorSet.Add(1, 255, 0, 0)
   
    ' Assign the color set to the line strip.
    oLineStrip.ColorSet = oColorSet
   
    ' The two spirals are currently on top of each other so translate the
    ' new one in the x direction so they're side by side.
    Dim oMatrix As Matrix
    Set oMatrix = oLineStripNode.Transformation
    Call oMatrix.SetTranslation(oTransGeom.CreateVector(15, 0, 0))
    oLineStripNode.Transformation = oMatrix
   
    ' Update the view to see the resulting spiral.
    oIG.UpdateOverlayGraphics ThisApplication.ActiveView
   
    'oIE.Stop
End Sub

