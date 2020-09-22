VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.MDIForm frmMain 
   Appearance      =   0  '¥­­±
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "ThincTalk"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   8745
   Icon            =   "MDIChat.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.PictureBox picSideBar 
      Align           =   4  '¹ï»ôªí³æ¥k¤è
      BackColor       =   &H00808080&
      BorderStyle     =   0  '¨S¦³®Ø½u
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6345
      Left            =   8610
      ScaleHeight     =   6345
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   0
      Width           =   135
      Begin VB.PictureBox PicTasks 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   6255
         Left            =   480
         ScaleHeight     =   6255
         ScaleWidth      =   2745
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   2745
         Begin VB.PictureBox tabPage 
            BorderStyle     =   0  '¨S¦³®Ø½u
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   0
            ScaleHeight     =   390
            ScaleMode       =   0  '¨Ï¥ÎªÌ¦Û­q
            ScaleWidth      =   2655
            TabIndex        =   13
            Top             =   0
            Width           =   2655
            Begin ThincTalk.chameleonButton btnTabButton 
               Height          =   360
               Index           =   0
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   635
               BTYPE           =   9
               TX              =   "News"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   0   'False
               BCOL            =   15133675
               BCOLO           =   15133675
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "MDIChat.frx":1982
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   -1  'True
               VALUE           =   -1  'True
            End
            Begin ThincTalk.chameleonButton btnTabButton 
               Height          =   360
               Index           =   1
               Left            =   360
               TabIndex        =   15
               Top             =   0
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   635
               BTYPE           =   9
               TX              =   "Contacts"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   0   'False
               BCOL            =   15133675
               BCOLO           =   15133675
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "MDIChat.frx":199E
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   -1  'True
               VALUE           =   0   'False
            End
            Begin ThincTalk.chameleonButton btnTabButton 
               Height          =   360
               Index           =   2
               Left            =   720
               TabIndex        =   16
               Top             =   0
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   635
               BTYPE           =   9
               TX              =   "Log"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   0   'False
               BCOL            =   15133675
               BCOLO           =   15133675
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "MDIChat.frx":19BA
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   -1  'True
               VALUE           =   0   'False
            End
            Begin ThincTalk.chameleonButton btnTabButton 
               Height          =   360
               Index           =   3
               Left            =   1080
               TabIndex        =   17
               Top             =   0
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   635
               BTYPE           =   9
               TX              =   "Info"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   0   'False
               BCOL            =   15133675
               BCOLO           =   15133675
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "MDIChat.frx":19D6
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   -1  'True
               VALUE           =   0   'False
            End
         End
         Begin VB.PictureBox picPage 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '¨S¦³®Ø½u
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Index           =   3
            Left            =   120
            ScaleHeight     =   4695
            ScaleWidth      =   2415
            TabIndex        =   12
            Top             =   480
            Width           =   2415
            Begin VB.Image imgDP 
               Appearance      =   0  '¥­­±
               BorderStyle     =   1  '³æ½u©T©w
               Height          =   1440
               Index           =   1
               Left            =   480
               Picture         =   "MDIChat.frx":19F2
               Stretch         =   -1  'True
               Top             =   360
               Width           =   1440
            End
            Begin VB.Image imgDP 
               Appearance      =   0  '¥­­±
               BorderStyle     =   1  '³æ½u©T©w
               Height          =   1440
               Index           =   0
               Left            =   480
               Picture         =   "MDIChat.frx":8634
               Stretch         =   -1  'True
               Top             =   3120
               Width           =   1440
            End
         End
         Begin VB.PictureBox picPage 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '¨S¦³®Ø½u
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Index           =   1
            Left            =   240
            ScaleHeight     =   4695
            ScaleWidth      =   2415
            TabIndex        =   4
            Top             =   360
            Width           =   2415
            Begin VB.ListBox lstContacts 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4140
               IntegralHeight  =   0   'False
               ItemData        =   "MDIChat.frx":F276
               Left            =   0
               List            =   "MDIChat.frx":F278
               TabIndex        =   8
               ToolTipText     =   "The list of contacts. Only works if the contact has a server open."
               Top             =   0
               Width           =   2280
            End
            Begin VB.ListBox lstBuddyIndex 
               Height          =   1680
               ItemData        =   "MDIChat.frx":F27A
               Left            =   720
               List            =   "MDIChat.frx":F27C
               TabIndex        =   7
               Top             =   600
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CommandButton btnDeleteContact 
               Caption         =   "&Delete"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1080
               TabIndex        =   6
               Top             =   3960
               Width           =   975
            End
            Begin VB.CommandButton btnAddContact 
               Caption         =   "&Add"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   5
               Top             =   3960
               Width           =   975
            End
         End
         Begin VB.PictureBox picPage 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '¨S¦³®Ø½u
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Index           =   2
            Left            =   120
            ScaleHeight     =   4695
            ScaleWidth      =   2415
            TabIndex        =   9
            Top             =   360
            Width           =   2415
            Begin VB.ListBox lstInfo 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1380
               IntegralHeight  =   0   'False
               ItemData        =   "MDIChat.frx":F27E
               Left            =   0
               List            =   "MDIChat.frx":F280
               TabIndex        =   11
               Top             =   0
               Width           =   2295
            End
            Begin VB.CommandButton btnClear 
               Caption         =   "&Clear"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   0
               TabIndex        =   10
               Top             =   4320
               Width           =   975
            End
            Begin VB.Label lblLines 
               AutoSize        =   -1  'True
               BackStyle       =   0  '³z©ú
               Caption         =   "0 Lines"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1080
               TabIndex        =   18
               Top             =   4380
               Width           =   495
            End
         End
         Begin VB.PictureBox picPage 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '¨S¦³®Ø½u
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Index           =   0
            Left            =   240
            ScaleHeight     =   4695
            ScaleWidth      =   2415
            TabIndex        =   2
            Top             =   360
            Width           =   2415
            Begin SHDocVwCtl.WebBrowser WB1 
               Height          =   975
               Left            =   0
               TabIndex        =   3
               Top             =   0
               Width           =   1095
               ExtentX         =   1931
               ExtentY         =   1720
               ViewMode        =   0
               Offline         =   0
               Silent          =   0
               RegisterAsBrowser=   0
               RegisterAsDropTarget=   1
               AutoArrange     =   0   'False
               NoClientEdge    =   0   'False
               AlignLeft       =   0   'False
               NoWebView       =   0   'False
               HideFileNames   =   0   'False
               SingleClick     =   0   'False
               SingleSelection =   0   'False
               NoFolders       =   0   'False
               Transparent     =   0   'False
               ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
               Location        =   "http:///"
            End
         End
      End
      Begin VB.Image Handle 
         Appearance      =   0  '¥­­±
         Height          =   1230
         Left            =   0
         MousePointer    =   9  'ªF-¦è¦V
         Picture         =   "MDIChat.frx":F282
         Top             =   2520
         Width           =   105
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewX As Long
Private Const BarWidth As Long = 2660



Private Sub btnTabButton_Click(Index As Integer)
    On Error Resume Next
    Dim A As Integer
    A = Index
    picPage(A).ZOrder 0
    picPage(A).Move 0, 315, tabPage.Width, PicTasks.Height - 315
    PicTasks_Resize
    For A = 0 To btnTabButton.UBound Step 1
        btnTabButton(A).Value = (A = Index)
    Next
    btnTabButton(WhichTab).Value = True
    For A = 0 To picPage.UBound
        picPage(A).Visible = (A = WhichTab)
    Next
    picPage(WhichTab).ZOrder 0
End Sub

Public Sub Handle_DblClick()
    On Error Resume Next
    picSideBar.Width = BarWidth
End Sub

Private Sub Handle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then NewX = X
End Sub

Private Sub Handle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        Dim A As Long
        'A = picSideBar.Width - NewX + X
        A = picSideBar.Width + NewX - X
        If A < Handle.Width + 300 Then A = Handle.Width   'Redraw
        If A > frmMain.Width - 300 Then A = frmMain.Width - Handle.Width
        If A < BarWidth + 300 And A > BarWidth - 300 Then A = BarWidth
        If A < frmMain.Width / 2 + 300 And A > frmMain.Width / 2 - 300 Then A = frmMain.Width / 2
        picSideBar.Width = A
        'Handle.Left = A - Handle.Width
        X = NewX
    End If
    PicTasks_Resize
End Sub

Public Function LoadBuddies()
    On Error Resume Next
    Dim i As Long
    Dim Buffer As String
    Log "Loading Buddies..."
    lstContacts.Clear
    lstBuddyIndex.Clear
    For i = 0 To MaxBuddies
        Buffer = FetchItem(i, 0)
        If Len(Buffer) > 0 Then
            lstContacts.AddItem Buffer
            'GetContactStatus I
            lstBuddyIndex.AddItem i
        End If
    Next
End Function
Private Sub btnAddContact_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    PopupMenu frmChat.titAddContact, , btnAddContact.Left, picPage(0).Top + btnAddContact.Top + btnAddContact.Height, frmChat.titaddthisperson
End Sub

Private Sub btnClear_Click()
    On Error Resume Next
    lstInfo.Clear
End Sub

Private Sub btnDeleteContact_Click()
On Error Resume Next
    If lstContacts.ListIndex < 0 Then Exit Sub
    If MyMsgbox(lstContacts.List(lstContacts.ListIndex) & vbCrLf & vbCrLf & "Are you sure you want to delete this contact?", _
    "DoNot5") = vbYes Then
        DelBuddy Val(lstBuddyIndex.List(lstContacts.ListIndex))
        LoadBuddies
    End If
End Sub

Private Sub lstInfo_DblClick()
    On Error Resume Next
    Clipboard.SetText lstInfo.List(lstInfo.ListIndex)
End Sub

Private Sub lstContacts_Click()
On Error Resume Next
    'GetContactStatus lstBuddyIndex.List(lstContacts.ListIndex)
    btnDeleteContact.Enabled = (lstContacts.ListIndex >= 0)
End Sub

Private Sub lstContacts_DblClick()
On Error Resume Next
    If MyMsgbox("This will end your current conversation, if you have one. Continue?", "DoNot6") = vbYes Then
        Sock1.Close
        EnableWindow False
        frmChat.StartClient FetchItem(Val(lstBuddyIndex.List(lstContacts.ListIndex)), 1), _
                        FetchItem(Val(lstBuddyIndex.List(lstContacts.ListIndex)), 2) 'starts new client
    End If
End Sub

Private Sub MDIForm_Load()
    On Error Resume Next
    btnTabButton_Click 0
    LoadBuddies
    Set A = New frmWizard
    A.Show
    WB1.Navigate2 GetSet("NewsURL", "http://www.kgv.net/blai/talk/news.htm", "Global")
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    Handle_MouseMove 1, 0, 0, 0
End Sub

Private Sub picPage_Resize(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0 'News
            WB1.Move 0, 0, picPage(Index).Width, picPage(Index).Height
        Case 1 'Contacts
            lstContacts.Move 0, 0, picPage(Index).Width, picPage(1).Height - btnAddContact.Height
            btnAddContact.Move 0, picPage(Index).Height - btnAddContact.Height
            btnDeleteContact.Move btnAddContact.Left + btnAddContact.Width, btnAddContact.Top
        Case 2 'Log
            btnClear.Move 0, picPage(Index).Height - btnClear.Height
            lstInfo.Move 0, 0, picPage(Index).Width, picPage(Index).Height - btnClear.Height
        Case 3 'Info
            imgDP(0).Move (picPage(Index).Width - imgDP(0).Width) / 2, picPage(3).Height - imgDP(0).Height - 360
            imgDP(1).Move (picPage(Index).Width - imgDP(1).Width) / 2, 360
    End Select
End Sub

Private Sub PicSideBar_Resize()
    On Error Resume Next
    Dim i As Integer
    PicTasks.Move Handle.Width, 0, picSideBar.Width - Handle.Width, Me.ScaleHeight
    'Handle.Move picSideBar.Width - Handle.Width, (picSideBar.Height - Handle.Height) / 2
    Handle.Move 0, (picSideBar.Height - Handle.Height) / 2
    picPage_Resize WhichTab
End Sub

Private Sub PicTasks_Resize()
    On Error Resume Next
    Dim i As Integer, A As Integer
    A = WhichTab
    tabPage.Move 0, 0, PicTasks.Width, PicTasks.Height
    picPage(A).Move 0, 315, tabPage.Width, PicTasks.Height - 315
    For i = 0 To picPage.UBound Step 1
        picPage(i).Visible = (i = A)
    Next
    picPage(A).Visible = (picPage(A).Width > 300)
    tabPage.Visible = (picPage(A).Width > 300)
    picPage_Resize A
End Sub

Private Sub TabPage_Resize()
    On Error Resume Next
    Dim i As Integer, j As Long
    For i = 0 To 3 Step 1
        With btnTabButton(i)
            .Left = j
            .Width = tabPage.Width / btnTabButton.Count - 15
            j = j + .Width + 15
            .Height = 315
        End With
    Next
End Sub

Private Sub WB1_NewWindow2(ppDisp As Object, Cancel As Boolean)
    On Error Resume Next
    Cancel = True
End Sub

Private Function WhichTab() As Integer
    On Error Resume Next
    Dim i As Integer
    For i = 0 To btnTabButton.UBound Step 1
        If btnTabButton(i).Value = True Then
            WhichTab = i
            Exit For
        End If
    Next
End Function
