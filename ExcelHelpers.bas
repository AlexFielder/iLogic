Attribute VB_Name = "ExcelHelpers"
Option Explicit

Function IsWorkbookOpen(ByVal Filename As String) As Boolean
    Dim TargetWorkbook As Workbook

    Dim IteratorWorkbook As Workbook
    For Each IteratorWorkbook In Application.Workbooks
        If IteratorWorkbook.FullName = Filename Then
            Set TargetWorkbook = IteratorWorkbook
        End If
    Next

    If Not TargetWorkbook Is Nothing Then
        If TargetWorkbook.ReadOnly Then
            IsWorkbookOpen = True
            Exit Function
        End If
    End If
End Function

'copied from here: https://stackoverflow.com/a/46028197/572634
'To export all the modules from 1 mother office file, to another .xlsm file:
'if the workbook that contains the mother macros is located in folder: "../a"
'Then place the child office files in: "../a/receiving/
'And create a(n empty) subfolder: "../a/receiving/modules
'Open the VBA for applications editor in MS Office, click "tools>references" and mark the checkbox: "Microsoft Scripting Runtime"


Sub Update_Workbooks()
'This macro requires that a reference to Microsoft Scripting Routine
'be selected under Tools\References in order for it to work.
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Dim fso As New FileSystemObject
Dim source As Scripting.Folder
Dim wbfile As Scripting.File
Dim book As Excel.Workbook
Dim sheet As Excel.Worksheet
Dim Filename As String
Dim ModuleFile As String
Dim Element As Object
Dim return_user_input As Integer


'Set source = fso.GetFolder("C:\Users\Desktop\Testing")   'we will know this since all of the files will be in one folder
Set source = fso.GetFolder(ThisWorkbook.path & "\receiving")   'we will know this since all of the files will be in one folder
For Each wbfile In source.Files

    Call basic_messagebox(wbfile, return_user_input)
    If return_user_input = 6 Then
        MsgBox (wbfile.Name & " yes")

        'For Each wbFile In source.Files

        If fso.GetExtensionName(wbfile.Name) = "xlsm" Then  'we will konw this too. All files will be .xlsm
            'Call basic_messagebox
            'Set book = Workbooks.Open(wbFile.Path)
            Set book = Workbooks.Open(wbfile.path)
                Filename = FileNameOnly(wbfile.Name)
                'This will remove all modules including ClassModules and UserForms.
                'It will keep all object modules like (sheets, ThisWorkbook)
                'On Error Resume Next
                Workbooks(wbfile.Name).Activate
                'For Each Element In ActiveWorkbook.VBProject.VBComponents
                'On Error Resume Next
                Call DeleteAllCode(wbfile)
'                For Each Element In Workbooks(wbfile.name).VBProject.VBComponents
'                    'ActiveWorkbook.VBProject.VBComponents.Remove Element
'                    Workbooks(wbfile.name).VBProject.VBComponents.Remove Element
'
'                Next
'                For Each Module In Workbooks(wbfile.name).VBProject.VBComponents
'                    Workbooks(wbfile.name).VBProject.VBComponents.Remove Module
'                Next

                'On Error GoTo ErrHandle
            '   Export Module1 from updating workbook
                'ModuleFile = Application.DefaultFilePath & "\tempmodxxx.bas"
                ModuleFile = ThisWorkbook.path & "\receiving\modules" & "\tempmodxxx.bas"

'                Workbooks("Update Multiple Workbooks.xlsm").VBProject.VBComponents("Module1") _
'                .Export ModuleFile
                'On Error Resume Next
                For Each Module In ThisWorkbook.VBProject.VBComponents
                    'MsgBox (Module.name)
                    If Left(Module.Name, 5) <> "Sheet" Then
                        If Left(Module.Name, 6) = "Module" Then
                            'MsgBox ("the modules name = " & Module.name)
                            'ThisWorkbook.VBProject.VBComponents("Module1").Export ModuleFile
                            ThisWorkbook.VBProject.VBComponents(Module.Name).Export ModuleFile
                            'ThisWorkbook.VBProject.VBComponents(Module).Export ModuleFile
                            'MsgBox (ModuleFile)
                        '   Replace Module1 in Userbook
                            Set VBP = Workbooks(Filename).VBProject
                            'On Error Resume Next
                            With VBP.VBComponents
                                .Import ModuleFile
                            End With
                        '   Delete the temporary module file
                            Kill ModuleFile
                        End If
                    End If
                Next
            'book.Close True
        End If
'Next


    End If
    If return_user_input = 7 Then
        MsgBox (wbfile.Name & " no")
    End If





Next

Exit Sub
ErrHandle:
'   Did an error occur?
    MsgBox "ERROR. The module may not have been replaced.", _
      vbCritical
End Sub

Private Function FileNameOnly(pname) As String
    Dim temp As Variant
    Length = Len(pname)
    temp = Split(pname, Application.PathSeparator)
    FileNameOnly = temp(UBound(temp))
End Function

Sub basic_messagebox(wbfile, return_user_input)
    'source: http://www.databison.com/vba-message-box-msgbox-the-message-can-do-better/
    'vbOK = 1
    'vbCancel = 2
    'vbAbort = 3
    'vbRetry = 4
    'vbIgnore = 5
    'vbYes = 6
    'vbNo = 7

    i = MsgBox("Do you wish to force the new code on the following excel file: " & vbNewLine & vbNewLine & wbfile.Name, vbYesNo)
    If i = 6 Then
        'MsgBox (wbFile.name & " yes")
        return_user_input = i
    End If
    If i = 7 Then
        'MsgBox (wbFile.name & " no")
        return_user_input = i
    End If

End Sub

Sub DeleteAllCode(wbfile)
     'Source: http://www.vbaexpress.com/kb/getarticle.php?kb_id=93
     'Trust Access To Visual Basics Project must be enabled.
     'From Excel: Tools | Macro | Security | Trusted Sources

    Dim x               As Integer
    Dim Proceed         As VbMsgBoxResult
    Dim Prompt          As String
    Dim Title           As String

    Prompt = "Are you certain that you want to delete all the VBA Code from " & _
    ActiveWorkbook.Name & "?"
    Title = "Verify Procedure"

    Proceed = MsgBox(Prompt, vbYesNo + vbQuestion, Title)
    If Proceed = vbNo Then
        MsgBox "Procedure Canceled", vbInformation, "Procedure Aborted"
        Exit Sub
    End If

    On Error Resume Next
    With ActiveWorkbook.VBProject
        For x = .VBComponents.count To 1 Step -1
            .VBComponents.Remove .VBComponents(x)
        Next x
        For x = .VBComponents.count To 1 Step -1
            .VBComponents(x).CodeModule.DeleteLines _
            1, .VBComponents(x).CodeModule.CountOfLines
        Next x
    End With
    On Error GoTo 0

End Sub
