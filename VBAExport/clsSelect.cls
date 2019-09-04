VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Print ' Declare the event objects
Private WithEvents oInteraction As InteractionEvents
Attribute oInteraction.VB_VarHelpID = -1
Private WithEvents oSelect As SelectEvents
Attribute oSelect.VB_VarHelpID = -1

' Declare a flag that's used to determine when selection stops.
Private bStillSelecting As Boolean

Public Function Pick(filter As SelectionFilterEnum) As Object
    ' Initialize flag.
    bStillSelecting = True
    
    ' Create an InteractionEvents object.
    Set oInteraction = ThisApplication.CommandManager.CreateInteractionEvents
    
    ' Define that we want select events rather than mouse events.
    oInteraction.SelectionActive = True

    ' Set a reference to the select events.
    Set oSelect = oInteraction.SelectEvents
    
    ' Set the filter using the value passed in.
    oSelect.AddSelectionFilter filter
    
    ' The InteractionEvents object.
    oInteraction.Start
    
    ' Loop until a selection is made.
    Do While bStillSelecting
        DoEvents
    Loop
    
    ' Get the selected item.  If more than one thing was selected,
    ' just get the first item and ignore the rest.
    Dim oSelectedEnts As ObjectsEnumerator
    Set oSelectedEnts = oSelect.SelectedEntities
    If oSelectedEnts.Count > 0 Then
        Set Pick = oSelectedEnts.Item(1)
    Else
        Set Pick = Nothing
    End If
    
    ' Stop the InteractionEvents object.
    oInteraction.Stop
    
    ' Clean up.
    Set oSelect = Nothing
    Set oInteraction = Nothing
End Function


Private Sub oInteraction_OnTerminate()
    ' Set the flag to indicate we're done.
    bStillSelecting = False
End Sub


Private Sub oSelect_OnSelect(ByVal JustSelectedEntities As ObjectsEnumerator, _
                                ByVal SelectionDevice As SelectionDeviceEnum, _
                                ByVal ModelPosition As Point, _
                                ByVal ViewPosition As Point2d, _
                                ByVal View As View)
    ' Set the flag to indicate we're done.
    bStillSelecting = False
End Sub

Public Function PickPartToMove(filter As SelectionFilterEnum) As ObjectsEnumerator
' Initialize flag.
    bStillSelecting = True
    
    ' Create an InteractionEvents object.
    Set oInteraction = ThisApplication.CommandManager.CreateInteractionEvents
    
    ' Define that we want select events rather than mouse events.
    oInteraction.SelectionActive = True

    ' Set a reference to the select events.
    Set oSelect = oInteraction.SelectEvents
    'oSelect.SingleSelectEnabled = False
    oSelect.WindowSelectEnabled = True
    ' Set the filter using the value passed in.
    oSelect.AddSelectionFilter filter

    ' The InteractionEvents object.
    oInteraction.Start
    
    ' Loop until a selection is made.
    Do While bStillSelecting
        DoEvents
    Loop
    
    ' Get the selected item.  If more than one thing was selected,
    ' just get the first item and ignore the rest.
    
    Dim oSelectedEnts As ObjectsEnumerator
    Set oSelectedEnts = oSelect.SelectedEntities
    
    If oSelectedEnts.Count > 0 Then
        Set PickPartToMove = oSelectedEnts
    Else
        Set PickPartToMove = Nothing
    End If
    
    ' Stop the InteractionEvents object.
    oInteraction.Stop
    
    ' Clean up.
    Set oSelect = Nothing
    Set oInteraction = Nothing
End Function
