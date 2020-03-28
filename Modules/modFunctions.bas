Attribute VB_Name = "modFunctions"
Option Explicit

Private Type POINTAPI
    X                                   As Long
    Y                                   As Long
End Type

Public Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function MessageBoxW Lib "user32.dll" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
Public Const WM_MOUSEMOVE               As Long = &H200
Public Const WM_LBUTTONDOWN             As Long = &H201
Private Const WM_LBUTTONUP              As Long = &H202
Private Const WM_MBUTTONDOWN            As Long = &H207
Private Const WM_MBUTTONUP              As Long = &H208
Private Const WM_RBUTTONDOWN            As Long = &H204
Private Const WM_RBUTTONUP              As Long = &H205
Private Const WM_NCLBUTTONDOWN          As Long = &HA1
Private Const WM_LBUTTONDBLCLK          As Long = &H203
Private Const WM_MBUTTONDBLCLK          As Long = &H209
Private Const WM_RBUTTONDBLCLK          As Long = &H206
Private Const MK_LBUTTON                As Long = &H1
Private Const MK_MBUTTON                As Long = &H10
Private Const MK_RBUTTON                As Long = &H2
Private Const WM_SETREDRAW              As Long = &HB
Private Const WM_ERASEBKGND             As Long = &H14
Private Const HWND_TOPMOST              As Long = -1
Private Const HWND_NOTOPMOST            As Long = -2
Private Const SWP_NOSIZE                As Long = &H1
Private Const SWP_NOMOVE                As Long = &H2

Public MousePoint                       As POINTAPI
Public sHwnd                            As Long
Public PointX                           As Integer
Public PointY                           As Integer

Public Function SetRedraw(ByVal sHwnd As Long, ByVal sBoolean As Boolean)
Dim ISend As Long
    ISend = SendMessage(sHwnd, WM_SETREDRAW, IIf(sBoolean, 1, 0), 0&)
End Function

Public Function MouseClick(ByVal sHwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal Index As Long) As Long
Dim lParam As Long
Dim ISend As Long
    lParam = (Y * &H10000) Or (X And &HFFFF&)
    Select Case Index
        Case 0 'Left Click
            ISend = PostMessage(sHwnd, WM_LBUTTONDOWN, MK_LBUTTON, ByVal lParam)
            ISend = PostMessage(sHwnd, WM_LBUTTONUP, 0, ByVal lParam)
        Case 1 'Left DblClick
            ISend = PostMessage(sHwnd, WM_LBUTTONDOWN, MK_LBUTTON, ByVal lParam)
            ISend = PostMessage(sHwnd, WM_LBUTTONUP, 0, ByVal lParam)
            ISend = PostMessage(sHwnd, WM_LBUTTONDBLCLK, MK_LBUTTON, ByVal lParam)
            ISend = PostMessage(sHwnd, WM_LBUTTONUP, 0, ByVal lParam)
        Case 2 'Middle Click
            ISend = PostMessage(sHwnd, WM_MBUTTONDOWN, MK_MBUTTON, ByVal lParam)
            ISend = PostMessage(sHwnd, WM_MBUTTONUP, 0, ByVal lParam)
        Case 3 'Middle DBlClick
            ISend = PostMessage(sHwnd, WM_MBUTTONDOWN, MK_MBUTTON, ByVal lParam)
            ISend = PostMessage(sHwnd, WM_MBUTTONUP, 0, ByVal lParam)
            ISend = PostMessage(sHwnd, WM_MBUTTONDBLCLK, MK_MBUTTON, ByVal lParam)
            ISend = PostMessage(sHwnd, WM_MBUTTONUP, 0, ByVal lParam)
        Case 4 'Right Click
            ISend = PostMessage(sHwnd, WM_RBUTTONDOWN, MK_RBUTTON, ByVal lParam)
            ISend = PostMessage(sHwnd, WM_RBUTTONUP, 0, ByVal lParam)
        Case 5 'Right DblClick
            ISend = PostMessage(sHwnd, WM_RBUTTONDOWN, MK_RBUTTON, ByVal lParam)
            ISend = PostMessage(sHwnd, WM_RBUTTONUP, 0, ByVal lParam)
            ISend = PostMessage(sHwnd, WM_RBUTTONDBLCLK, MK_RBUTTON, ByVal lParam)
            ISend = PostMessage(sHwnd, WM_RBUTTONUP, 0, ByVal lParam)
    End Select
End Function

Public Function OnTop(hwnd As Long, Value As Boolean) As Long
    If Value = True Then
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Function

Public Function UniMsgBox(strText As String, Optional iButtons As VbMsgBoxStyle = vbOKOnly, Optional strTitle As String, Optional hwnd As Long = &H0) As VbMsgBoxResult
    UniMsgBox = MessageBoxW(hwnd, StrPtr(ToUni(strText)), StrPtr(ToUni(strTitle)), iButtons)
End Function

Private Function ToUni(str$) As String
    Dim ansi$, UNI$, i&, sTem$, sUni$, arrUNI() As String
    ansi = "a1|a2|a3|a4|a5|a6|a8|a61a62a63a64a65a81a82a83a84a85A1|A2|A3|A4|A5|A6|A8|A61A62A63A64A65A81A82A83A84A85e1|e2|e3|e4|e5|e6|e61e62e63e64e65E1|E2|E3|E4|E5|E6|E61E62E63E64E65i1|i2|i3|i4|i5|I1|I2|I3|I4|I5|o1|o2|o3|o4|o5|o6|o7|o61o62o63o64o65o71o72o73o74o75O1|O2|O3|O4|O5|O6|O7|O61O62O63O64O65O71O72O73O74O75u1|u2|u3|u4|u5|u7|u71u72u73u74u75U1|U2|U3|U4|U5|U7|U71U72U73U74U75y1|y2|y3|y4|y5|Y1|Y2|Y3|Y4|Y5|d9|D9|"
    UNI = "E1,E0,1EA3,E3,1EA1,E2,103,1EA5,1EA7,1EA9,1EAB,1EAD,1EAF,1EB1,1EB3,1EB5,1EB7,C1,C0,1EA2,C3,1EA0,C2,102,1EA4,1EA6,1EA8,1EAA,1EAC,1EAE,1EB0,1EB2,1EB4,1EB6,E9,E8,1EBB,1EBD,1EB9,EA,1EBF,1EC1,1EC3,1EC5,1EC7,C9,C8,1EBA,1EBC,1EB8,CA,1EBE,1EC0,1EC2,1EC4,1EC6,ED,EC,1EC9,129,1ECB,CD,CC,1EC8,128,1ECA,F3,F2,1ECF,F5,1ECD,F4,1A1,1ED1,1ED3,1ED5,1ED7,1ED9,1EDB,1EDD,1EDF,1EE1,1EE3,D3,D2,1ECE,D5,1ECC,D4,1A0,1ED0,1ED2,1ED4,1ED6,1ED8,1EDA,1EDC,1EDE,1EE0,1EE2,FA,F9,1EE7,169,1EE5,1B0,1EE9,1EEB,1EED,1EEF,1EF1,DA,D9,1EE6,168,1EE4,1AF,1EE8,1EEA,1EEC,1EEE,1EF0,FD,1EF3,1EF7,1EF9,1EF5,DD,1EF2,1EF6,1EF8,1EF4,111,110"
    arrUNI = Split(UNI, ",")

    For i = 1 To Len(str)
        If IsNumeric(Mid$(str, i + 1, 1)) = False Then
            sUni = sUni & Mid$(str, i, 1)
        Else
            sTem = IIf(IsNumeric(Mid$(str, i + 2, 1)), Mid$(str, i, 3), Mid$(str, i, 2))
            i = i + IIf(IsNumeric(Mid$(str, i + 2, 1)), 2, 1)
            If InStr(ansi, sTem) > 0 Then sTem = ChrW("&H" & arrUNI(InStr(ansi, sTem) \ 3))
            sUni = sUni & sTem
        End If
    Next
    ToUni = sUni
End Function

