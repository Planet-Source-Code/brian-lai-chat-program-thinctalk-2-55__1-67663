VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Form"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   -360
      Width           =   375
   End
   Begin VB.CommandButton btnEnd 
      Caption         =   "&No"
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton btnEnd 
      Caption         =   "&Yes"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      Width           =   855
   End
   Begin VB.CheckBox chkDoNotShowAgain 
      Caption         =   "Remember my &Answer"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1260
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frmMsgBox.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Label1"
      Height          =   1050
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   3405
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEnd_Click(Index As Integer)
    On Error Resume Next
    ReturnValue = Index
    If chkDoNotShowAgain.Value = 1 Then 'If to remember
        SaveSet Me.Tag, Str(ReturnValue), lblMsg.Tag
    End If
    Text1.Text = ""
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Caption = App.ProductName
    SkinForm Me
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyC And Shift = 2 Then
        Clipboard.SetText lblMsg.Caption
    ElseIf KeyCode = vbKeyY Then
        btnEnd_Click 1
    ElseIf KeyCode = vbKeyN Then
        btnEnd_Click 2
    ElseIf KeyCode = vbKeyReturn Then
        btnEnd_Click 1
    Else
        Text1.Text = ""
    End If
End Sub
