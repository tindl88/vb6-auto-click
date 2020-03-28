Attribute VB_Name = "modMain"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Sub Main()
    On Error Resume Next
    InitCommonControls
    frmMain.Show
End Sub

