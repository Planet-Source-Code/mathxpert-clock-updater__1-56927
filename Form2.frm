VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Clock"
   ClientHeight    =   240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4560
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngTimer As Long

Private Sub Form_GotFocus()
If bFocus And Form1.Visible Then Form1.SetFocus
End Sub

Private Sub Form_Load()
Left = Form1.Left
Top = Form1.Top + Form1.Height
Label1.Caption = "System time:  " & Format$(Now, "M/D/YY  H:MM:SS." & MS & " AM/PM")
lngTimer = SetTimer(0, 0, 1, AddressOf TimerProc)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
KillTimer 0, lngTimer
If bUncheck Then
    If Not bUnloading Then Form1.Check1.Value = 0
End If
End Sub
