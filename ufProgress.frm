VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufProgress 
   Caption         =   "Running..."
   ClientHeight    =   1620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ufProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PLACE IN YOUR USERFORM CODE
' copied from here: https://wellsr.com/vba/2017/excel/beautiful-vba-progress-bar-with-step-by-step-instructions/
Private Sub UserForm_Initialize()
#If IsMac = False Then
    'hide the title bar if you're working on a windows machine. Otherwise, just display it as you normally would
    Me.Height = Me.Height - 10
    HideTitleBar.HideTitleBar Me
#End If
End Sub
