Attribute VB_Name = "modSysTray"
Option Explicit

Private Const MAX_TOOLTIP   As Integer = 64
Private Const NIF_ICON      As Long = &H2
Private Const NIF_MESSAGE   As Long = &H1
Private Const NIF_TIP       As Long = &H4
Private Const NIM_ADD       As Long = &H0
Private Const NIM_DELETE    As Long = &H2
'Private Const WM_MOUSEMOVE   As Long = &H200
'Private Const WM_LBUTTONDOWN As Long = &H201

Private Type NOTIFYICONDATA
    cbSize                  As Long
    hwnd                    As Long
    uID                     As Long
    uFlags                  As Long
    uCallbackMessage        As Long
    hIcon                   As Long
    szTip                   As String * MAX_TOOLTIP
End Type
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private nfIconData          As NOTIFYICONDATA

Public Sub AddTrayIcon(Frm As Form, ByVal sText As String)
    With nfIconData
        .hwnd = Frm.hwnd
        .uID = Frm.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Frm.Icon.Handle
        .szTip = sText & vbNullChar
        .cbSize = Len(nfIconData)
    End With
    Call Shell_NotifyIcon(NIM_ADD, nfIconData)
End Sub

Public Sub RemoveTrayIcon(Frm As Form)
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub
