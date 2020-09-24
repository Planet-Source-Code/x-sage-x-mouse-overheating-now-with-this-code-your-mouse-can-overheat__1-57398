VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mouse Overheating"
   ClientHeight    =   285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleWidth      =   1710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   1665
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   105
         TabIndex        =   1
         Top             =   0
         Width           =   135
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Dim Val As Integer
Dim OldX As Integer
Dim OldY As Integer
Dim aX As Integer
Dim aY As Integer
Dim Posit As POINTAPI
Dim Over As Integer
Private Sub Form_Load()
Val = 0
Over = 10
End Sub

Private Sub Timer1_Timer()
If Over = 0 Then
Call GetCursorPos(Posit)
If Val > Picture1.Width Then
Call SetCursorPos(OldX, OldY)
Over = 100
Else
aX = Posit.X
aY = Posit.Y
End If

If Val > Picture1.Width / 3 * 2 Then
Picture2.BackColor = vbRed
ElseIf Val > Picture1.Width / 3 Then
Picture2.BackColor = vbYellow
ElseIf Val < Picture1.Width / 3 Then
Picture2.BackColor = vbGreen
End If

Val = Val + Abs(aX - OldX) + Abs(aY - OldY)
If Val >= 10 Then
Val = Val - 10
End If
Picture2.Width = Val
OldX = aX
OldY = aY
End If
If Over > 0 Then
Picture2.BackColor = vbBlue
Call SetCursorPos(OldX, OldY)
Over = Over - 1
End If
End Sub
