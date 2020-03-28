VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto Click"
   ClientHeight    =   1365
   ClientLeft      =   6540
   ClientTop       =   4155
   ClientWidth     =   2880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkOnTop 
      Alignment       =   1  'Right Justify
      Caption         =   "On Top"
      Height          =   195
      Left            =   1980
      TabIndex        =   5
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   405
      Left            =   1680
      TabIndex        =   3
      Top             =   900
      Width           =   1155
   End
   Begin VB.ComboBox cboMouse 
      Height          =   315
      ItemData        =   "frmMain.frx":0E42
      Left            =   60
      List            =   "frmMain.frx":0E44
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   540
      Width           =   1635
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   405
      Left            =   60
      TabIndex        =   0
      Top             =   900
      Width           =   1605
   End
   Begin VB.TextBox txtDelay 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2250
      TabIndex        =   1
      Text            =   "1000"
      Top             =   540
      Width           =   435
   End
   Begin VB.Timer tmrAutoClick 
      Enabled         =   0   'False
      Left            =   2430
      Top             =   30
   End
   Begin VB.Image imgTarget 
      Height          =   480
      Left            =   1200
      Top             =   30
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Delay:           s"
      Height          =   195
      Left            =   1770
      TabIndex        =   4
      Top             =   600
      Width           =   1035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkOnTop_Click()
    Call OnTop(Me.hwnd, chkOnTop.Value)
End Sub

Private Sub tmrAutoClick_Timer()
    MouseClick sHwnd, PointX, PointY, Val(cboMouse.ListIndex)
End Sub

Private Sub cmdStart_Click()
    If sHwnd = 0 Then
        UniMsgBox "Ba5n ha4y kéo hình vuông màu d9o3 to71i vi5 trí ca62n Click." & vbCrLf & _
                  "    - Chu7o7ng trình không chie61m chuo65t." & vbCrLf & _
                  "    - Có the63 di chuye63n và thu nho3 (Minimize) d9o61i tu7o75ng ca62n click." & vbCrLf & _
                  "    - Tho72i gian click (Delay) nho3 nha61t la2 0.1 = 100 mili giây, lo71n nha61t là 60 giây." & vbCrLf & vbCrLf & _
                  "Các ba5n có the63 ta3i ma4 nguo62n cu3a chu7o7ng trình ta5i website: www.caulacbovb.com." & vbCrLf & _
                  "Ta1c gia3: tindl88@yahoo.com", vbInformation, "Clicker", Me.hwnd
        Exit Sub
    End If
    If cmdStart.Tag = 0 Then
        If Val(txtDelay.Text) > 60 Then txtDelay.Text = 60
        If txtDelay.Text = "" Then txtDelay.Text = "1"
        MouseClick sHwnd, PointX, PointY, Val(cboMouse.ListIndex)
        tmrAutoClick.Interval = Val(txtDelay.Text) * 1000
        tmrAutoClick.Enabled = True
        cmdStart.Caption = "Stop"
        cboMouse.Enabled = False
        txtDelay.Enabled = False
        imgTarget.Enabled = False
        cmdStart.Tag = 1
    Else
        tmrAutoClick.Enabled = False
        cmdStart.Caption = "Start"
        cboMouse.Enabled = True
        txtDelay.Enabled = True
        imgTarget.Enabled = True
        cmdStart.Tag = 0
    End If
End Sub

Private Sub imgTarget_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgTarget.Picture = Nothing
    Me.MousePointer = 99
    Me.MouseIcon = Me.Icon
End Sub

Private Sub imgTarget_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgTarget.Picture = Me.Icon
    Me.MousePointer = 0
    Call GetCursorPos(MousePoint)
    sHwnd = WindowFromPoint(MousePoint.X, MousePoint.Y)
    Call ScreenToClient(sHwnd, MousePoint)
    PointX = MousePoint.X
    PointY = MousePoint.Y
End Sub

Private Sub Form_Load()
On Error Resume Next
    cboMouse.AddItem "Left Click"
    cboMouse.AddItem "Left DblClick"
    cboMouse.AddItem "Middle Click"
    cboMouse.AddItem "Middle DblClick"
    cboMouse.AddItem "Right Click"
    cboMouse.AddItem "Right DblClick"
    cmdStart.Tag = 0
    Me.Tag = 0
    imgTarget.Picture = Me.Icon
    AddTrayIcon Me, Me.Caption & Space$(1) & App.Major & "." & App.Minor
    cboMouse.ListIndex = GetSetting("AutoClick", "Setting", "Mouse?", 0)
    txtDelay.Text = GetSetting("AutoClick", "Setting", "Delay?", "3.1")
    Me.Left = GetSetting("AutoClick", "Setting", "Left?", (Screen.Width - Me.Width) / 2)
    Me.Top = GetSetting("AutoClick", "Setting", "Top?", (Screen.Height - Me.Height) / 2)
    If Me.Tag = GetSetting("AutoClick", "Setting", "First?", 0) Then UniMsgBox "Ba5n ha4y kéo hình vuông màu d9o3 to71i vi5 trí ca62n Click." & vbCrLf & "    - Chu7o7ng trình không chie61m chuo65t." & vbCrLf & "    - Có the63 di chuye63n và thu nho3 (Minimize) d9o61i tu7o75ng ca62n click." & vbCrLf & "    - Tho72i gian click (Delay) nho3 nha61t la2 0.1 = 100 mili giây, lo71n nha61t là 60 giây." & vbCrLf & vbCrLf & "Ta1c gia3: tindl88@yahoo.com", vbInformation, "Clicker", Me.hwnd
    chkOnTop.Value = GetSetting("AutoClick", "Setting", "OnTop?", 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RemoveTrayIcon Me
    SaveSetting "AutoClick", "Setting", "Mouse?", cboMouse.ListIndex
    SaveSetting "AutoClick", "Setting", "Delay?", txtDelay.Text
    SaveSetting "AutoClick", "Setting", "Left?", Me.Left
    SaveSetting "AutoClick", "Setting", "Top?", Me.Top
    SaveSetting "AutoClick", "Setting", "First?", 1
    SaveSetting "AutoClick", "Setting", "OnTop?", chkOnTop.Value
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lMsg As Single
    lMsg = X / Screen.TwipsPerPixelX
    Select Case lMsg
        Case WM_LBUTTONDOWN: Me.Show
    End Select
End Sub

Private Sub txtDelay_KeyPress(KeyAscii As Integer)
    If InStr("1234567890." + Chr$(vbKeyBack), Chr$(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub cmdHide_Click()
    Me.Hide
End Sub
