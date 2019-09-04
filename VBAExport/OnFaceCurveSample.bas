Attribute VB_Name = "OnFaceCurveSample"
Sub OnFaceCurveSample()
    Dim oDoc As PartDocument
    Set oDoc = ThisApplication.ActiveDocument
    
    Dim oFace As Face
    Set oFace = oDoc.ComponentDefinition.SurfaceBodies(1).Faces(1)
     
    Dim oSk3D As Sketch3D
    Set oSk3D = oDoc.ComponentDefinition.Sketches3D.Add
    
    Dim oFaces As NameValueMap
    Dim oFitPoints As NameValueMap
    
    Set oFaces = ThisApplication.TransientObjects.CreateNameValueMap
    Set oFitPoints = ThisApplication.TransientObjects.CreateNameValueMap
    
    Dim i As Long, oTempFace As Face, oCol As ObjectCollection
    For i = 1 To oFace.FaceShell.Faces.Count
        Set oCol = ThisApplication.TransientObjects.CreateObjectCollection
        
        Set oTempFace = oFace.FaceShell.Faces(i)
        oCol.Add oTempFace.PointOnFace
        
        oFaces.Add "Face" & i, oTempFace
        oFitPoints.Add "Face" & i, oCol
    Next

    Dim oOnFaceCurve  As OnFaceCurve

    Set oOnFaceCurve = oSk3D.OnFaceCurves.Add(oFaces, oFitPoints)
End Sub
