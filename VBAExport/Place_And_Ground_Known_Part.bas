Attribute VB_Name = "Place_And_Ground_Known_Part"
Sub Main()
Dim filename As String
filename = "C:\Users\Alex.Fielder\OneDrive\Inventor\Designs\C360\Part5.ipt"
Call Place_and_Ground_Part(ThisApplication, filename)
End Sub

Public Function Place_and_Ground_Part(ByVal invApp As Application, ByVal path As String) As ComponentOccurrence

    ' Post the filename to the private event queue.
    Call invApp.CommandManager.PostPrivateEvent(Inventor.PrivateEventTypeEnum.kFileNameEvent, path)

    ' Get the control definition for the Place Component command.
    Dim ctrlDef As Inventor.ControlDefinition
    Set ctrlDef = invApp.CommandManager.ControlDefinitions.Item("AssemblyPlaceComponentCmd")

    ' Execute the command.
    Call ctrlDef.Execute

    'Return Nothing
End Function
