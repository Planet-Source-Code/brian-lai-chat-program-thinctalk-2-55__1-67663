VERSION 5.00
Begin VB.Form frmAddContact 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Add a Contact"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtUserData 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox txtUserData 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox txtUserData 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  '¸m¤¤¹ï»ô
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Add Contact to Contact List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   3915
   End
   Begin VB.Label lblUserData 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "When you add a contact here, you will be able to connect to the server they opened if they keep the server open all the time."
      Height          =   585
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   3765
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblUserData 
      BackStyle       =   0  '³z©ú
      Caption         =   "Contact's Server Port:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label lblUserData 
      BackStyle       =   0  '³z©ú
      Caption         =   "Contact's Server IP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label lblUserData 
      BackStyle       =   0  '³z©ú
      Caption         =   "Contact Nickname:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3735
   End
   Begin VB.Image imgThisBkg 
      Height          =   675
      Left            =   0
      Picture         =   "frmAddContact.frx":1982
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmAddContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To 2 Step 1
        If Len(txtUserData(i).Text) = 0 Then
            MsgBox "One of the fields is empty.", vbExclamation
            Exit Sub
        End If
    Next
    If Len(txtUserData(2).Text) = 0 Or txtUserData(2).Text = "0" Then txtUserData(1).Text = DefaultPort
    AddBuddy txtUserData(0).Text, txtUserData(1).Text, txtUserData(2).Text
    frmMain.LoadBuddies
    Unload Me
End Sub

Public Function LoadAddCotnacts(Optional Who As String, Optional Where As String, Optional Which As Long)
    On Error Resume Next
    txtUserData(0).Text = Who
    txtUserData(1).Text = Where
    txtUserData(2).Text = Str(Which)
    Me.Show 1
End Function
