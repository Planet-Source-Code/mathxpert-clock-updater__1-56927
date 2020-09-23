VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clock Updater"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Tag             =   "Clock Updater"
   Visible         =   0   'False
   Begin VB.CheckBox Check2 
      Caption         =   "Run at Startup"
      Height          =   255
      Left            =   2423
      TabIndex        =   17
      Top             =   2520
      Width           =   1328
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View Clock"
      Height          =   255
      Left            =   1103
      TabIndex        =   16
      Top             =   2520
      Width           =   1088
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":1CFA
      Left            =   1103
      List            =   "Form1.frx":1D28
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3315
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop synchro"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2348
      TabIndex        =   19
      ToolTipText     =   "Stop adjustment"
      Top             =   2880
      Width           =   2085
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   3023
      Max             =   12
      TabIndex        =   14
      Top             =   2160
      Value           =   11
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Synchronize now"
      Height          =   375
      Left            =   128
      TabIndex        =   18
      ToolTipText     =   "Set PC clock to atomic standard"
      Top             =   2880
      Width           =   2085
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Options:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   915
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000005&
      Height          =   180
      Left            =   3958
      Top             =   765
      Width           =   450
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000005&
      Height          =   180
      Left            =   3488
      Top             =   765
      Width           =   450
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000005&
      Height          =   180
      Left            =   3018
      Top             =   765
      Width           =   450
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000005&
      Height          =   180
      Left            =   2548
      Top             =   765
      Width           =   450
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Time server:"
      Height          =   255
      Left            =   143
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   480
      Picture         =   "Form1.frx":1E6B
      Top             =   3480
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "Form1.frx":1FB5
      Top             =   3480
      Width           =   240
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1103
      TabIndex        =   11
      Top             =   1800
      Width           =   3315
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Next update:"
      Height          =   255
      Left            =   143
      TabIndex        =   10
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Left            =   2663
      TabIndex        =   13
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000005&
      Height          =   180
      Left            =   1138
      Top             =   765
      Width           =   450
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000005&
      Height          =   180
      Left            =   1608
      Top             =   765
      Width           =   450
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000005&
      Height          =   180
      Left            =   2078
      Top             =   765
      Width           =   450
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1103
      TabIndex        =   9
      Top             =   1440
      Width           =   3315
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1103
      TabIndex        =   7
      Top             =   1080
      Width           =   3315
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Time after:"
      Height          =   255
      Left            =   143
      TabIndex        =   8
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Time before:"
      Height          =   255
      Left            =   143
      TabIndex        =   6
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1103
      TabIndex        =   4
      Top             =   480
      Width           =   3315
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1103
      TabIndex        =   5
      Top             =   720
      Width           =   3315
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Sync status:"
      Height          =   255
      Left            =   143
      TabIndex        =   3
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label8 
      Caption         =   "Sync period (in hours) (0 = manual):"
      Height          =   255
      Left            =   143
      TabIndex        =   12
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Menu mPopUpSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Show"
      End
      Begin VB.Menu mPopDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mPopSync 
         Caption         =   "&Sync now"
      End
      Begin VB.Menu mPopStop 
         Caption         =   "&Stop sync"
         Enabled         =   0   'False
      End
      Begin VB.Menu mPopDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChangeIt As Boolean

Public Function Retrieve(sData As String, Optional DefData) As String
If IsMissing(DefData) Then
    Retrieve = GetSetting(App.Title, "Data", sData)
Else
    Retrieve = GetSetting(App.Title, "Data", sData, DefData)
End If
End Function

Public Function RetrieveNum(sTitle As String, sData As String, Optional DefData)
Dim MyMsg
If IsMissing(DefData) Then
    MyMsg = GetSetting(App.Title, sTitle, sData)
Else
    MyMsg = GetSetting(App.Title, sTitle, sData, DefData)
End If
If IsNumeric(MyMsg) Then
    RetrieveNum = MyMsg
Else
    If IsMissing(DefData) Then
        RetrieveNum = 0
    Else
        RetrieveNum = DefData
    End If
End If
End Function

Public Sub Save(sData As String, Data)
SaveSetting App.Title, "Data", sData, Data
End Sub

Public Sub Save2(sTitle As String, sData As String, Data)
SaveSetting App.Title, sTitle, sData, Data
End Sub

Public Sub ChangeIcon(Optional RestoreBack As Boolean = False)
With mtIconData
    .hIcon = IIf(RestoreBack, Image2.Picture, Image1.Picture)
    mtIconData.uFlags = NIF_ICON
End With
Shell_NotifyIcon NIM_MODIFY, mtIconData
End Sub

Public Sub ChangeInterval(ByVal HourlyInterval As Long)
If HourlyInterval <> 0 Then
    RunTimer = SetTimer(0&, 0&, HourlyInterval * 36000000, AddressOf OnSynchro)
    Label10.Caption = Format(DateAdd("H", HourlyInterval, Now), "M/D/YY  H:MM:SS." & MS & " AM/PM")
End If
TimeD = HourlyInterval
End Sub

Private Sub Check1_Click()
If bFocus Then Text1.SetFocus

If Check1.Value = 1 Then
    If Not bUncheck Then bUncheck = True
    Form2.Show , Me
Else
    If bUncheck Then bUncheck = False
    Unload Form2
End If
End Sub

Private Sub Check2_Click()
If bFocus Then Text1.SetFocus
End Sub

Private Sub Combo1_Click()
If bFocus Then Text1.SetFocus
End Sub

Private Sub Command1_Click()
If bFocus Then Text1.SetFocus
OnSynchro
End Sub

Private Sub Command2_Click()
If bFocus Then Text1.SetFocus
Cancelled = True
End Sub

Private Sub Form_Load()
Dim st As String, cmVal As Long, cmVal2 As Long, cmVal3 As Integer, cmVal4 As String

bUsed = False
bFocus = False
bUncheck = True
bUnloading = False

If App.PrevInstance Then End

cmVal = RetrieveNum("Servers", "ServerSetting", 0)
If cmVal >= 0 And cmVal <= Combo1.ListCount Then
    Combo1.ListIndex = cmVal
Else
    Combo1.ListIndex = 0
End If

cmVal2 = RetrieveNum("HowOften", "Value", 1)
If cmVal2 >= 0 And cmVal2 <= 12 Then
    VScroll1.Value = 12 - cmVal2
Else
    VScroll1.Value = 11
End If

cmVal3 = RetrieveNum("Misc", "ClockDisp", 0)
If cmVal3 = 0 Or cmVal3 = 1 Then
    Check1.Value = cmVal3
Else
    Check1.Value = 0
End If

If ValueExists(HKEY_CURRENT_USER, GetStartupRegValue, "ClockUpdater") Then
    cmVal4 = GetRegValue(HKEY_CURRENT_USER, GetStartupRegValue, "ClockUpdater")
    If cmVal4 <> FormatAppPath(App.Path & "\" & App.EXEName & ".exe") Then
        SetRegValue HKEY_CURRENT_USER, GetStartupRegValue, "ClockUpdater", FormatAppPath(App.Path & "\" & App.EXEName & ".exe")
    End If
    Check2.Value = 1
Else
    Check2.Value = 0
End If

Label4.Caption = Retrieve("Before")
Label5.Caption = Retrieve("After")

Show
With mtIconData
    .cbSize = Len(mtIconData)
    .hwnd = Me.hwnd
    .uCallbackMessage = WM_MOUSEMOVE
    .uID = 1&
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .hIcon = Image2.Picture
    .szTip = Me.Tag & Chr$(0)
    If Shell_NotifyIcon(NIM_ADD, mtIconData) = 0 Then
        MsgBox "Failed to create tray icon!", vbExclamation, "Cannot create icon"
        End
    End If
End With
WindowState = 1

st = Retrieve("Status")
bFocus = True
Label1.Caption = IIf(st = "", "Ready", st)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Local Error Resume Next
Err.Clear

Static bBusy As Boolean
    If Not bBusy Then
        bBusy = True
        Select Case CLng(x) / 15
            Case WM_LBUTTONDBLCLK
                Me.WindowState = 0
                Me.Visible = True
                DoEvents
                AppActivate "Clock Updater"
                If bFocus Then Text1.SetFocus
                
            Case WM_LBUTTONDOWN
                
            Case WM_LBUTTONUP
                
            Case WM_RBUTTONDBLCLK
                
            Case WM_RBUTTONDOWN
            
            Case WM_RBUTTONUP
                With mPopRestore
                    If Me.Visible Then
                        If .Caption = "&Show" Then .Caption = "&Hide"
                        If ChangeIt = False Then ChangeIt = True
                    Else
                        If .Caption = "&Hide" Then .Caption = "&Show"
                        If ChangeIt Then ChangeIt = False
                    End If
                End With
                Me.PopupMenu Me.mPopUpSys, , , , mPopRestore
                
        End Select
        bBusy = False
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
bUnloading = True
End Sub

Private Sub Form_Resize()
If WindowState = 1 Then Me.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Save "Status", IIf(Label1.Caption = "Synchronizing; please wait...", "Synchronization aborted at " & Format(Now, "M/D/YY  H:MM:SS." & MS & " AM/PM"), Label1.Caption)
Save "Before", Label4.Caption
Save "After", Label5.Caption
Save2 "Servers", "ServerSetting", Combo1.ListIndex
Save2 "HowOften", "Value", 12 - VScroll1.Value
Save2 "Misc", "ClockDisp", Check1.Value

If Check2.Value = 1 Then
    If Not ValueExists(HKEY_CURRENT_USER, GetStartupRegValue, "ClockUpdater") Then SetRegValue HKEY_CURRENT_USER, GetStartupRegValue, "ClockUpdater", FormatAppPath(App.Path & "\" & App.EXEName & ".exe")
Else
    If ValueExists(HKEY_CURRENT_USER, GetStartupRegValue, "ClockUpdater") Then DeleteValue HKEY_CURRENT_USER, GetStartupRegValue, "ClockUpdater"
End If

Unload Form2
KillTimer 0&, RunTimer
Shell_NotifyIcon NIM_DELETE, mtIconData
End Sub

Private Sub mPopExit_Click()
Unload Me
End Sub

Private Sub mPopRestore_Click()
If ChangeIt = True Then
    WindowState = 1
Else
    Me.WindowState = 0
    Me.Visible = True
    DoEvents
    AppActivate "Clock Updater"
End If
End Sub

Private Sub mPopStop_Click()
Cancelled = True
End Sub

Private Sub mPopSync_Click()
OnSynchro
End Sub

Private Sub VScroll1_Change()
If bFocus Then Text1.SetFocus
Select Case VScroll1.Value
    Case 12
        Label7.Caption = 0
        KillTimer 0&, RunTimer
        Label10.Caption = ""
    Case 11
        Label7.Caption = 1
        ChangeInterval 1
    Case 10
        Label7.Caption = 2
        ChangeInterval 2
    Case 9
        Label7.Caption = 3
        ChangeInterval 3
    Case 8
        Label7.Caption = 4
        ChangeInterval 4
    Case 7
        Label7.Caption = 5
        ChangeInterval 5
    Case 6
        Label7.Caption = 6
        ChangeInterval 6
    Case 5
        Label7.Caption = 7
        ChangeInterval 7
    Case 4
        Label7.Caption = 8
        ChangeInterval 8
    Case 3
        Label7.Caption = 9
        ChangeInterval 9
    Case 2
        Label7.Caption = 10
        ChangeInterval 10
    Case 1
        Label7.Caption = 11
        ChangeInterval 11
    Case 0
        Label7.Caption = 12
        ChangeInterval 12
End Select
End Sub
