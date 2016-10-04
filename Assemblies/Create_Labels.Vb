Imports Microsoft.VisualBasic
Imports System.IO
Imports Microsoft.Office.Interop.Excel

Public Class Class1
    ''' <summary>
    ''' The thinking behind this iLogic script is that rather than manually create a series of labels controlled by a single detail drawing,
    ''' that we use a combination of Excel and iLogic to create the files for us.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Main()
        Call createLabels()
    End Sub
    Public Sub createLabels()
        'first we need to open the excel file
        Dim NewFilesArray As New ArrayList
        Dim
        'this can be forgotten once we use this in iLogic
        '''
        Dim excel As application = New application
        Dim wb As Workbook = excel.workbooks.open("")
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = wb.Worksheets(1)
        Dim r As Range = ws.UsedRange
        '''

        NewFilesArray = GoExcel.CellValues("C:\VAULT WORKING FOLDER\Designs\2011 Designs\AWESOME\Models\DT-999999-000.xlsx")

        For MyRow = 2 To 1000
            ' column A thru E
            If Not CStr(GoExcel.CellValue("A" & MyRow) = "") Then
                Dim lbl As New Label
                lbl.PartNo = GoExcel.CellValue("A" & MyRow)
                lbl.NewLabelText = GoExcel.CellValue("B" & MyRow)
                lbl.NewCharHeight = GoExcel.CellValue("C" & MyRow)
                lbl.NewDimA = GoExcel.CellValue("D" & MyRow)
                lbl.NewDimB = GoExcel.CellValue("E" & MyRow)
                Call CreateNewLabel(lbl)
            End If
        Next

    End Sub
    Public Sub CreateNewLabel(newlabel As Label)

    End Sub
    Private Function returnColumnName(columnNo As Integer) As String
        Dim columnName As String
        Select Case columnNo
            Case 1
                columnName = "A"
            Case 2
                columnName = "B"
            Case 3
                columnName = "C"
            Case 4
                columnName = "D"
            Case 5
                columnName = "E"
            Case 6
                columnName = "F"
            Case 7
                columnName = "G"
            Case 8
                columnName = "H"
            Case 9
                columnName = "I"
            Case 10
                columnName = "J"
            Case 11
                columnName = "K"
            Case 12
                columnName = "L"
            Case 13
                columnName = "M"
            Case 14
                columnName = "N"
            Case 15
                columnName = "O"
            Case 16
                columnName = "P"
            Case 17
                columnName = "Q"
            Case 18
                columnName = "R"
            Case 19
                columnName = "S"
            Case 20
                columnname = "T"
            Case 21
                columnname = "U"
            Case 22
                columnname = "V"
            Case 23
                columnname = "W"
            Case 24
                columnname = "X"
            Case 25
                columnname = "Y"
            Case 26
                columnname = "Z"
        End Select
        Return columnName
    End Function


End Class

Public Class Label
    
    Public PartNo As [String]
    
    Public NewLabelText As [String]
    
    Public NewCharHeight As [String]

    Public NewDimA As [Double]

    Public NewDimB As [Double]

    Public Sub [New](m_partno As [String],
                     m_labeltext As [String],
                     m_charheight As [Double],
                     m_dima As [Double],
                     m_dimb As [Double])
        PartNo = m_partno
        NewLabelText = m_labeltext
        NewCharHeight = m_charheight
        NewDimA = m_dima
        NewDimB = m_dimb
    End Sub



End Class

