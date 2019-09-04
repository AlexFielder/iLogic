Attribute VB_Name = "DebugExcelRangeObject"
Sub DebugSelectedCell()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim ws As Worksheet
    Set ws = wb.ActiveSheet
    Dim cell As Range
    Set cell = ActiveCell
    Stop
End Sub
