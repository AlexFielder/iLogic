Attribute VB_Name = "BomTools"
Public Sub BOMQuery()
    ' Set a reference to the assembly document.
    ' This assumes an assembly document is active.
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument

    Dim FirstLevelOnly As Boolean
    If MsgBox("First level only?", vbYesNo) = vbYes Then
        FirstLevelOnly = True
    Else
        FirstLevelOnly = False
    End If
    
    ' Set a reference to the BOM
    Dim oBOM As BOM
    Set oBOM = oDoc.ComponentDefinition.BOM
    
    ' Set whether first level only or all levels.
    If FirstLevelOnly Then
        oBOM.StructuredViewFirstLevelOnly = True
    Else
        oBOM.StructuredViewFirstLevelOnly = False
    End If
    
    ' Make sure that the structured view is enabled.
    oBOM.StructuredViewEnabled = True
    
    'Set a reference to the "Structured" BOMView
    Dim oBOMView As BOMView
    Set oBOMView = oBOM.BOMViews.Item("Structured")
        
    Debug.Print "Item"; Tab(15); "Quantity"; Tab(30); "Part Number"; Tab(70); "Description"
    Debug.Print "----------------------------------------------------------------------------------"

    'Initialize the tab for ItemNumber
    Dim ItemTab As Long
    ItemTab = -3
    Call QueryBOMRowProperties(oBOMView.BOMRows, ItemTab)
End Sub

Private Sub QueryBOMRowProperties(oBOMRows As BOMRowsEnumerator, ItemTab As Long)
    ItemTab = ItemTab + 3
    ' Iterate through the contents of the BOM Rows.
    Dim i As Long
    For i = 1 To oBOMRows.Count
        ' Get the current row.
        Dim oRow As BOMRow
        Set oRow = oBOMRows.Item(i)

        'Set a reference to the primary ComponentDefinition of the row
        Dim oCompDef As ComponentDefinition
        Set oCompDef = oRow.ComponentDefinitions.Item(1)

        Dim oPartNumProperty As Property
        Dim oDescripProperty As Property

        If TypeOf oCompDef Is VirtualComponentDefinition Then
            'Get the file property that contains the "Part Number"
            'The file property is obtained from the virtual component definition
            Set oPartNumProperty = oCompDef.PropertySets _
                .Item("Design Tracking Properties").Item("Part Number")

            'Get the file property that contains the "Description"
            Set oDescripProperty = oCompDef.PropertySets _
                .Item("Design Tracking Properties").Item("Description")

            Debug.Print Tab(ItemTab); oRow.ItemNumber; Tab(17); oRow.ItemQuantity; Tab(30); _
                oPartNumProperty.Value; Tab(70); oDescripProperty.Value
        Else
            'Get the file property that contains the "Part Number"
            'The file property is obtained from the parent
            'document of the associated ComponentDefinition.
            Set oPartNumProperty = oCompDef.Document.PropertySets _
                .Item("Design Tracking Properties").Item("Part Number")

            'Get the file property that contains the "Description"
            Set oDescripProperty = oCompDef.Document.PropertySets _
                .Item("Design Tracking Properties").Item("Description")

            Debug.Print Tab(ItemTab); oRow.ItemNumber; Tab(17); oRow.ItemQuantity; Tab(30); _
                oPartNumProperty.Value; Tab(70); oDescripProperty.Value
            
            'Recursively iterate child rows if present.
            If Not oRow.ChildRows Is Nothing Then
                Call QueryBOMRowProperties(oRow.ChildRows, ItemTab)
            End If
        End If
    Next
    ItemTab = ItemTab - 3
End Sub

Public Sub BOMExport()
    ' Set a reference to the assembly document.
    ' This assumes an assembly document is active.
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument

    ' Set a reference to the BOM
    Dim oBOM As BOM
    Set oBOM = oDoc.ComponentDefinition.BOM
    
    ' Set the structured view to 'all levels'
    oBOM.StructuredViewFirstLevelOnly = False

    ' Make sure that the structured view is enabled.
    oBOM.StructuredViewEnabled = True

    ' Set a reference to the "Structured" BOMView
    Dim oStructuredBOMView As BOMView
    Set oStructuredBOMView = oBOM.BOMViews.Item("Structured")
    
    ' Export the BOM view to an Excel file
    oStructuredBOMView.Export "C:\temp\BOM-StructuredAllLevels.xls", kMicrosoftExcelFormat
  
    ' Make sure that the parts only view is enabled.
    oBOM.PartsOnlyViewEnabled = True

    ' Set a reference to the "Parts Only" BOMView
    Dim oPartsOnlyBOMView As BOMView
    Set oPartsOnlyBOMView = oBOM.BOMViews.Item("Parts Only")

    ' Export the BOM view to an Excel file
    oPartsOnlyBOMView.Export "C:\temp\BOM-PartsOnly.xls", kMicrosoftExcelFormat
End Sub



