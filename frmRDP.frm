VERSION 5.00
Begin VB.Form frmRDP 
   Caption         =   "Remote Desktop"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5775
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   5775
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.PictureBox PIcDesk 
      BackColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4275
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmRDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NewX As Long, NewY As Long

Private Sub Form_Activate()
    On Error Resume Next
    Me.Icon = frmMain.Icon
    If GetSet("OpenTwip", "0") = "1" Then SendDataEx "SckCMD:OpenTwip"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PIcDesk.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub PIcDesk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then 'Left
        Call SendDataEx("SckCMD:mouse 2,0,0")
    ElseIf Button = 2 Then 'Right
        Call SendDataEx("SckCMD:mouse 8,0,0")
    ElseIf Button = 4 Then 'Middle
        Call SendDataEx("SckCMD:mouse 16,0,0")
    End If
    NewX = X
    NewY = Y
End Sub

Private Sub PIcDesk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TehX As Long, TehY As Long
    TehX = -(NewX - X) / Screen.TwipsPerPixelX
    TehY = -(NewY - Y) / Screen.TwipsPerPixelX
    Call SendDataEx("SckCMD:mouse 1," & TehX & "," & TehY) '& ",0,0")
    NewX = X
    NewY = Y
End Sub

Private Sub PIcDesk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then 'Left
        Call SendDataEx("SckCMD:mouse 4,0,0")
    ElseIf Button = 2 Then 'Right
        Call SendDataEx("SckCMD:mouse 16,0,0")
    ElseIf Button = 4 Then 'Middle
        Call SendDataEx("SckCMD:mouse 64,0,0")
    End If
End Sub
