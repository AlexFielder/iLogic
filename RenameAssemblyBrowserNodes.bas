Attribute VB_Name = "RenameAssemblyBrowserNodes"
'credit goes to this thread: https://forums.autodesk.com/t5/inventor-customization/using-ilogic-to-rename-browser-nodes/m-p/4318636#M44814
' and PaulM for pointing me towards it.
Sub Main()
    'Grab the Assembly Document
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
'   Dim trans As Transaction = ThisApplication.TransactionManager.StartTransaction(oDoc, "Rename Browser Nodes to Description.")
'   Try
        Dim oAsmCompDef As AssemblyComponentDefinition
        Set oAsmCompDef = oDoc.ComponentDefinition
        Dim oPane As BrowserPane
        Set oPane = oDoc.BrowserPanes.Item("Model")
        Dim oOcc As ComponentOccurrence
        For Each oOcc In oAsmCompDef.Occurrences
            Dim oCCDocument As Document
            Set oCCDocument = oOcc.Definition.Document
            If Not Contains(oCCDocument.FullFileName, "Content Center") And Not oOcc.IsSubstituteOccurrence Then
                MsgBox (oOcc.name)
                Dim invDesignInfo As PropertySet
                Set invDesignInfo = oCCDocument.PropertySets.Item("Design Tracking Properties")
                Dim invDescrProperty As Inventor.Property
                Set invDescrProperty = invDesignInfo.Item("Description")
                Dim oSubAssyNode As BrowserNode
                Set oSubAssyNode = oPane.GetBrowserNodeFromObject(oOcc)
                If Contains(oSubAssyNode.NativeObject.name, ":") Then ' is likely one of multiple occurrences
                    'messagebox.Show(oSubAssyNode.NativeObject.Name)
                    Dim oldName As String
                    oldName = oSubAssyNode.NativeObject.name
                    Dim first As Integer
                    first = InStr(oldName, ":") 'oldName.InStr(":")
                    'MsgBox ("first= " & first)
                    Dim last As Integer
                    last = InStrRev(oldName, ":") 'oldName.LastIndexOf(":")
                    'MessageBox.Show("last= " & last)
                    Dim occNum As String
                    'occNum = oldName.Substring(first, Len(oldName) - last)
                    'Messagebox.Show(occNum)
                    If Not oldName = invDescrProperty.Value And Not invDescrProperty.Value = "" Then
                        'Set The name
                        oSubAssyNode.NativeObject.name = (invDescrProperty.Value) & occNum
                    End If
                Else
                    If oSubAssyNode.NativeObject.name <> invDescrProperty.Value And Not invDescrProperty.Value = "" Then
                    'Set The name
                    oSubAssyNode.NativeObject.name = (invDescrProperty.Value)
                    End If
                End If

            End If
        Next
'   Catch Ex As Exception
'       trans.Abort()
'   Finally
'       trans.End()
'   End Try
End Sub

Public Function Contains(strBaseString As String, strSearchTerm As String) As Boolean
    'Purpose: Returns TRUE if one string exists within another
    On Error GoTo ErrorMessage
        Contains = InStr(strBaseString, strSearchTerm)
    Exit Function
ErrorMessage:
    MsgBox "The database has generated an error. Please contact the database administrator, quoting the following error message: '" & Err.Description & "'", vbCritical, "Database Error"
    End
End Function
