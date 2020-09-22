VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmChat 
   Caption         =   "Hi"
   ClientHeight    =   5490
   ClientLeft      =   240
   ClientTop       =   4830
   ClientWidth     =   7305
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
   Icon            =   "frmMain1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   7305
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.PictureBox picInfoBar 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      BackColor       =   &H00000000&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7305
      TabIndex        =   10
      Top             =   0
      Width           =   7305
      Begin VB.Image imgDP 
         Appearance      =   0  '¥­­±
         BorderStyle     =   1  '³æ½u©T©w
         Height          =   480
         Index           =   1
         Left            =   120
         Stretch         =   -1  'True
         ToolTipText     =   "Your Display Picture"
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblContactInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "My Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   720
         TabIndex        =   16
         ToolTipText     =   "Your Nickname"
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lblContactInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "My IP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   15
         ToolTipText     =   "Your IP"
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblContactInfo 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Contact IP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   3360
         TabIndex        =   12
         ToolTipText     =   "Contact IP"
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblContactInfo 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Contact Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   2880
         TabIndex        =   11
         ToolTipText     =   "Contact name"
         Top             =   120
         Width           =   1380
      End
      Begin VB.Image imgDP 
         Appearance      =   0  '¥­­±
         BorderStyle     =   1  '³æ½u©T©w
         Height          =   480
         Index           =   0
         Left            =   4320
         Picture         =   "frmMain1.frx":000C
         Stretch         =   -1  'True
         ToolTipText     =   "Contact's Display Picture"
         Top             =   120
         Width           =   480
      End
      Begin VB.Image imgThisBkg 
         Height          =   675
         Left            =   0
         Picture         =   "frmMain1.frx":0B16
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox PicConvo 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '¨S¦³®Ø½u
      FillColor       =   &H8000000B&
      ForeColor       =   &H8000000F&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3975
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   840
      Width           =   4335
      Begin VB.PictureBox TB1 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   390
         Left            =   0
         ScaleHeight     =   390
         ScaleWidth      =   3375
         TabIndex        =   4
         Top             =   240
         Width           =   3375
         Begin ThincTalk.chameleonButton btnTB 
            Height          =   360
            Index           =   0
            Left            =   0
            TabIndex        =   5
            ToolTipText     =   "Font"
            Top             =   15
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            BTYPE           =   9
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   8421504
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16711935
            MPTR            =   1
            MICON           =   "frmMain1.frx":0C0C
            PICN            =   "frmMain1.frx":0C28
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ThincTalk.chameleonButton btnTB 
            Height          =   360
            Index           =   1
            Left            =   1680
            TabIndex        =   6
            ToolTipText     =   "Interaction"
            Top             =   15
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            BTYPE           =   9
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   8421504
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16711935
            MPTR            =   1
            MICON           =   "frmMain1.frx":0F7A
            PICN            =   "frmMain1.frx":0F96
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin ThincTalk.chameleonButton btnTB 
            Height          =   360
            Index           =   2
            Left            =   840
            TabIndex        =   7
            ToolTipText     =   "Encryption"
            Top             =   15
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            BTYPE           =   9
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   8421504
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16711935
            MPTR            =   1
            MICON           =   "frmMain1.frx":11F2
            PICN            =   "frmMain1.frx":120E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin ThincTalk.chameleonButton btnTB 
            Height          =   360
            Index           =   3
            Left            =   1215
            TabIndex        =   8
            ToolTipText     =   "Leet"
            Top             =   15
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            BTYPE           =   9
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   8421504
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16711935
            MPTR            =   1
            MICON           =   "frmMain1.frx":1560
            PICN            =   "frmMain1.frx":157C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin ThincTalk.chameleonButton btnTB 
            Height          =   360
            Index           =   5
            Left            =   2055
            TabIndex        =   9
            ToolTipText     =   "Remote Desktop (not done)"
            Top             =   15
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            BTYPE           =   9
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   8421504
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16711935
            MPTR            =   1
            MICON           =   "frmMain1.frx":18CE
            PICN            =   "frmMain1.frx":18EA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin ThincTalk.chameleonButton btnTB 
            Height          =   360
            Index           =   4
            Left            =   375
            TabIndex        =   13
            ToolTipText     =   "Contact info"
            Top             =   15
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            BTYPE           =   9
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   8421504
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   16711935
            MPTR            =   1
            MICON           =   "frmMain1.frx":1C3C
            PICN            =   "frmMain1.frx":1C58
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.CommandButton btnOK 
         Caption         =   "&Send"
         Default         =   -1  'True
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   3120
         Width           =   735
      End
      Begin VB.ComboBox txtSendMsg 
         Height          =   315
         Left            =   0
         TabIndex        =   1
         Top             =   3360
         Width           =   3495
      End
      Begin SHDocVwCtl.WebBrowser txtConvo 
         Height          =   2895
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   4215
         ExtentX         =   7435
         ExtentY         =   5106
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
      Begin VB.Line lineConvo 
         Visible         =   0   'False
         X1              =   0
         X2              =   4320
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  '³z©ú
         Caption         =   "Ready"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   3720
         UseMnemonic     =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Timer TimerRoutines 
      Interval        =   5000
      Left            =   6720
      Top             =   4920
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4680
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSWinsockLib.Winsock Sock3 
      Left            =   6120
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Sock2 
      Left            =   5640
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Sock1 
      Left            =   5160
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu titFile 
      Caption         =   "&File"
      Begin VB.Menu titAutoMSG 
         Caption         =   "Auto Message..."
         Checked         =   -1  'True
      End
      Begin VB.Menu titSendAny 
         Caption         =   "Send Command..."
         Shortcut        =   +{F8}
      End
      Begin VB.Menu titSaveLog 
         Caption         =   "Save Log..."
         Shortcut        =   ^S
      End
      Begin VB.Menu titAddContact 
         Caption         =   "Save to contact list..."
         Begin VB.Menu titaddthisperson 
            Caption         =   "This person"
         End
         Begin VB.Menu titAddContactEx 
            Caption         =   "Someone else..."
         End
      End
      Begin VB.Menu titN 
         Caption         =   "-"
      End
      Begin VB.Menu titPrefs 
         Caption         =   "Prefs..."
         Shortcut        =   ^R
      End
      Begin VB.Menu titN6 
         Caption         =   "-"
      End
      Begin VB.Menu titClose 
         Caption         =   "Close"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu titTalk 
      Caption         =   "&View"
      Begin VB.Menu titContactList 
         Caption         =   "Contact List"
      End
      Begin VB.Menu titChgNN 
         Caption         =   "Change Nickname"
      End
      Begin VB.Menu titTrans 
         Caption         =   "Opacity"
         Begin VB.Menu titTransparency 
            Caption         =   "10%"
            Index           =   1
         End
         Begin VB.Menu titTransparency 
            Caption         =   "20%"
            Index           =   2
         End
         Begin VB.Menu titTransparency 
            Caption         =   "30%"
            Index           =   3
         End
         Begin VB.Menu titTransparency 
            Caption         =   "40%"
            Index           =   4
         End
         Begin VB.Menu titTransparency 
            Caption         =   "50%"
            Index           =   5
         End
         Begin VB.Menu titTransparency 
            Caption         =   "60%"
            Index           =   6
         End
         Begin VB.Menu titTransparency 
            Caption         =   "70%"
            Index           =   7
         End
         Begin VB.Menu titTransparency 
            Caption         =   "80%"
            Index           =   8
         End
         Begin VB.Menu titTransparency 
            Caption         =   "90%"
            Index           =   9
         End
         Begin VB.Menu titTransparency 
            Caption         =   "Opaque"
            Index           =   10
         End
      End
   End
   Begin VB.Menu titSend 
      Caption         =   "Send"
      Begin VB.Menu titSendFile 
         Caption         =   "File..."
         Enabled         =   0   'False
      End
      Begin VB.Menu titSendImg 
         Caption         =   "Image..."
      End
      Begin VB.Menu titSendFlash 
         Caption         =   "Flash..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu titWindows 
      Caption         =   "Windows"
      WindowList      =   -1  'True
   End
   Begin VB.Menu PopupMenuForLabels1 
      Caption         =   "PopupMenuForLabels1"
      Visible         =   0   'False
      Begin VB.Menu titCopyCaption 
         Caption         =   "Copy"
      End
   End
   Begin VB.Menu PopupMenuForDP1 
      Caption         =   "PopupMenuForDP1"
      Visible         =   0   'False
      Begin VB.Menu titChangeDP 
         Caption         =   "Change Picture"
      End
      Begin VB.Menu titRefreshPic 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu PopupMenuForDP2 
      Caption         =   "PopupMenuForDP2"
      Visible         =   0   'False
      Begin VB.Menu titRefreshPicEx 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Voice As New SpVoice
Public IsServerMode As Boolean, TypingModeSent As Boolean
Public ServerPort As Long, ClientPort As Long
Public NickName As String, ClientNick As String, LastPersonTalking As String, StrAutoMSG As String, ClientVer As String
Dim LastSendAutoMSG As Double
Dim IsReady As Boolean
Private Const NoNickDieRound As Long = 400

Private Sub btnOK_Click()
On Error GoTo Er
    Dim q As String
    If Len(txtSendMsg.Text) = 0 Then Exit Sub
    If ClientNick = "" Then SendDataEx "SckGNk:", True
    q = txtSendMsg.Text
    If GetSet("StripHTML", "1") = "1" Then 'Remove HTML tags so users cant mess with them
        q = StripHTML(q, "")
    End If
    txtSendMsg.AddItem q, 0 'Add to 0, so it appears first on the list
    If Left$(q, 1) = "/" Then 'If it's a command (command routine at bottom)
        Dim MyCmd As String
        MyCmd = Mid$(q, 2) 'excluding the /
        SendDataEx "SckCMD:" & MyCmd, True 'Send the command then?
    Else 'If it's not a command
        If GetSet("LeetData", "0") = "1" Then q = RunTranslate(q)
'        Q = EncodeEmoCodes(Q) 'Make encoded emo codes
        If GetSet("AddName", "1") = "1" Then 'adding names stuff
            Dim ReplStr As String 'Dim for replacing nickname with style
            ReplStr = Replace(GetSet("NickStyle", "[::]"), "::", NickName) 'This is the transformed style
            Say ReplStr & " " & q, 0 'DISPLAY ONLY
        Else
            Say q, 0 'DISPLAY ONLY
        End If
        If GetSet("SpeakSend", "0") = "1" Then
            Voice.Speak q, SVSFlagsAsync 'Speak!
        End If
        SendData q, False, (GetSet("EncryptData", "1") = 1) 'SEND
        SendDataEx "SckTyp:0"
    End If
'    If GetSet("RemoveLineBreaks", "0") = "1" Then RemoveLineBreaks txtConvo ' Remove Extra vbCrs
    txtSendMsg.Text = ""
    txtSendMsg.SetFocus
    TypingModeSent = False
    btnOK.Enabled = IIf(GetSet("ForceSend", "0") = "1", True, False)
    
    Exit Sub
Er:
    AddText "Sending Failed. " & ClientNick & " did not get your message."
End Sub

Private Sub btnTB_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0 'Colour
            OpenFont 0
            ComboFont
            Form_Resize
            SendFont 0
            Form_Resize
        Case 1 'Interaction
            SaveSet "AllowRemoteExecution", Abs(btnTB(Index).Value)
        Case 2 'Encryption
            If btnTB(Index).Value = False Then
                If MyMsgbox("It is always good to encrypt the sent data so nobody else will understand it." & _
                vbCrLf & vbCrLf & "Do you really want to disable encryption?", "DoNot4") = vbNo Then
                    btnTB(Index).Value = True
                    Exit Sub
                End If
            End If
            AddText "Encryption has been turned " & IIf(btnTB(Index).Value = True, "on", "off"), True
            SendData NickName & " turned " & IIf(btnTB(Index).Value = True, "on", "off") & _
                            " his/her text encryption so whatever he/she sends you will " & _
                            IIf(btnTB(Index).Value = False, "not ", "") & "be encrypted.", True, GetSet("EncryptData", "1")
            SaveSet "EncryptData", Abs(btnTB(Index).Value)
        Case 3 '1337
            SaveSet "LeetData", Abs(btnTB(Index).Value)
            AddText "Leet Coder has been turned " & IIf(btnTB(Index).Value = True, "on", "off"), True
        Case 4 'Info
            If Len(Sock1.RemoteHostIP) > 0 Then
                ShowContactinfo
            Else
                MsgBox "You don't seem to be connected to a contact." & vbCrLf & _
                "You need to chat with another person to use this feature.", vbCritical
            End If
        Case 5 'RDP
            If btnTB(Index).Value = True Then
                frmRDP.Show
            Else
                Unload frmRDP
            End If
    End Select
    btnTB(5).Enabled = btnTB(1).Value 'Remote is available only if interaction is on
    Call SaveSet("FontName", FontNme(0))
    Call SaveSet("FontBold", Str(FontBld(0)))
    Call SaveSet("FontItalic", Str(FontItl(0)))
    Call SaveSet("FontUnderline", Str(FontUdl(0)))
    Call SaveSet("FontSize", Str(FontSze(0)))
    Call SaveSet("FontColor", Str(FontClr(0)))
    Call SaveSet("FontStrike", Str(FontStr(0)))
    ComboFont
    SendFont 0
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    txtConvo.Silent = True
    txtConvo.Navigate2 FindPath(App.Path, "tmppage.html")
End Sub

Private Sub Form_Load()
On Error Resume Next
    UpdateCaption
    SStatus
    titAutoMSG.Checked = False
    txtConvo.BorderStyle = IIf(GetSet("ConvoBorder", "1") = "1", rtfFixedSingle, rtfNoBorder)
    EnableWindow False
    FontNme(0) = GetSet("FontName", "Tahoma") 'load your own fonts
    FontSze(0) = Val(GetSet("FontSize", "10"))
    FontBld(0) = CBool(GetSet("FontBold", "False"))
    FontItl(0) = CBool(GetSet("FontItalic", "False"))
    FontUdl(0) = GetSet("FontUnderline", "False")
    FontStr(0) = CBool(GetSet("FontStrike", "False"))
    FontClr(0) = GetSet("FontColor", "0")
    Call ComboFont
    Call UpdateTB
    'DockingStart Me, True
    SkinForm Me
    Form_Resize
    IsReady = True
End Sub

Public Function AddEmos()
    Dim FoundPos As Long
    Dim K As Long
    For K = 200 To 300 Step 1 'Range for emos
        FoundPos = txtConvo.Find(LoadResString(i), 0, , rtfNoHighlight)
        While FoundPos > 0
            With txtConvo
                .SelStart = FoundPos
                .SelLength = 6
                .SelText = ""
                '''
                Debug.Print App.Path & "\emos\smile" & Expand(Str(K)) & ".bmp"
                .OLEObjects.Add , , App.Path & "\emos\smile" & Expand(Str(K)) & ".bmp" 'Add the picture after it has deleted the string
                '''
                DoEvents
                FoundPos = txtConvo.Find(LoadResString(i), FoundPos + 6, , rtfNoHighlight)
            End With
        Wend
    Next
End Function

Public Function EncodeEmoCodes(WhatHere As String) As String
    On Error Resume Next
    Dim i As Long
    For i = 200 To 300 Step 1
        WhatHere = Replace(Replace(WhatHere, LoadResString(i), "tbl" & Expand(Str$(i))), " ", "")
    Next
    EncodeEmoCodes = WhatHere
End Function

Public Function DecodeEmoCodes(WhatHere As String) As String
    On Error Resume Next
    Dim i As Long
    For i = 200 To 300 Step 1
        WhatHere = Replace(Replace(WhatHere, "tbl" & Expand(Str$(i)), LoadResString(i)), " ", "")
    Next
    DecodeEmoCodes = WhatHere
End Function

Public Function UpdateTB()
    On Error Resume Next
    btnTB(1).Value = IIf(GetSet("AllowRemoteExecution", "0") = "1", True, False)
    btnTB(2).Value = IIf(GetSet("EncryptData", "1") = "1", True, False)
    btnTB(3).Value = IIf(GetSet("LeetData", "0") = "1", True, False)
    btnTB(5).Enabled = btnTB(1).Value
End Function

Public Function ComboFont()
    On Error Resume Next
    With txtSendMsg
        .FontName = FontNme(0)
        .FontBold = FontBld(0)
        .FontItalic = FontItl(0)
        .FontUnderline = FontUdl(0)
        .FontSize = FontSze(0)
        .FontStrikethru = FontStr(0)
        .ForeColor = FontClr(0)
    End With
End Function

Public Function UpdateCaption()
On Error Resume Next
    Dim A As String, B As String
    frmMain.Caption = App.ProductName & IIf(IsServerMode = True, " Server", "")
    Me.Caption = NickName & IIf(Len(ClientNick) > 0, " with " & ClientNick, "")
    lblContactInfo(0).Caption = IIf(Len(ClientNick) > 0, ClientNick, "Not Connected")
    lblContactInfo(1).Caption = IIf(Len(Sock1.RemoteHostIP) > 0, Sock1.RemoteHostIP, "No IP")
    lblContactInfo(2).Caption = Sock1.LocalIP
    lblContactInfo(3).Caption = NickName
    A = GetSet("DisplayPic", "http://www.kgv.net/blai/Images/manshead.bmp")
    If Len(A) > 0 Then
        If imgDP(1).Tag <> A Then
            ChangeDP A
        End If
    End If
    
End Function

Public Function StartServer(ListenToPort As Long)
On Error GoTo Er
    IsServerMode = True
    ServerPort = ListenToPort
    Sock1.Close
    Sock1.LocalPort = ListenToPort
    If Left$(Sock1.LocalIP, 7) = "192.168" Then
        AddText "WARNING: this is a LAN IP, that means only people in your local area network can connect to you."
    End If
    Sock1.Listen
    'Availability Server
    Sock2.Close
    Sock2.LocalPort = AvailabilityPort
    Sock2.Listen
    '/Availability
    AddText "Waiting for contacts... (" & Sock1.LocalIP & ":" & ListenToPort & ")"
    Form_Load
    Exit Function
Er:
    AddText "Failed to Listen to port " & ListenToPort & ". " & vbCrLf & "Please make sure other applications are not using port " & _
    ListenToPort & ", or change your settings to use some other ports."
End Function

Public Function StartClient(ServerIP As String, ServerPort As Long)
On Error GoTo Er
    Dim PortSuccess As Boolean
    If ServerPort < 1 Or ServerPort > 65536 Then
        If GetSet("BlockErroneousPorts", "1") = "1" Then
            AddText "The port number is impossible for connecting."
            PortSuccess = False
        Else
            PortSuccess = True 'Ignore Port Error
        End If
    Else
        PortSuccess = True
        AddText "Please wait while " & App.ProductName & " connects to " & ServerIP & "."
    End If
    IsServerMode = False
    If ServerPort = 99999 Then ServerPort = DefaultPort
    ClientPort = ServerPort 'ServerPort here is a sub-only variable
    Sock1.Close
    Form_Load 'switched places. was after "Wait".
    If PortSuccess = True Then Sock1.Connect ServerIP, ServerPort 'connect only if the port is possible
    Wait
    Exit Function
Er:
    AddText "Failed to connect to Server " & ServerIP & _
    ". Maybe the server is behind a router or the server is using some other ports."
End Function

Private Sub Form_Resize()
On Error Resume Next
    PicConvo.Move 0, picInfoBar.Height, Me.ScaleWidth - PicConvo.Left, Me.ScaleHeight - picInfoBar.Height
    txtConvo.Move 0, 0, PicConvo.Width, PicConvo.Height - (txtSendMsg.Height + lblStatus.Height + TB1.Height) ' - 15 - txtConvo.Top
    TB1.Move 0, txtConvo.Top + txtConvo.Height, PicConvo.Width
    txtSendMsg.Move 0, TB1.Top + TB1.Height, PicConvo.Width - btnOK.Width
    lblStatus.Move 0, txtSendMsg.Top + txtSendMsg.Height, txtSendMsg.Width
    btnOK.Move txtSendMsg.Left + txtSendMsg.Width, TB1.Top + TB1.Height, btnOK.Width, txtSendMsg.Height + lblStatus.Height
    imgDP(0).Left = Me.ScaleWidth - imgDP(0).Width - 120
    lblContactInfo(0).Left = imgDP(0).Left - lblContactInfo(0).Width - 120
    lblContactInfo(1).Left = imgDP(0).Left - lblContactInfo(1).Width - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Dim dx As Form
    For Each dx In Forms
        Unload dx
    Next
'    End
End Sub

Private Sub imgDP_Click(Index As Integer)
    If Index = 1 Then 'ME
        PopupMenu PopupMenuForDP1, , imgDP(1).Left, imgDP(1).Top + imgDP(1).Height, titChangeDP
    ElseIf Index = 0 Then 'contact
        PopupMenu PopupMenuForDP2, , imgDP(0).Left, imgDP(0).Top + imgDP(0).Height, titRefreshPicEx
    End If
End Sub

Private Sub lblContactInfo_Click(Index As Integer)
    lblContactInfo(0).Tag = Index 'Temp Storage
    PopupMenu PopupMenuForLabels1, , lblContactInfo(Index).Left, _
                    lblContactInfo(Index).Top + lblContactInfo(Index).Height, titCopyCaption
    lblContactInfo(0).Tag = "" 'Popupmenu ends AFTER the menu disappears, so this var is useless now
End Sub

Private Sub picInfoBar_Resize()
    On Error Resume Next
    imgThisBkg.Move 0, 0, picInfoBar.Width, picInfoBar.Height
End Sub

Private Sub Sock1_Close()
    On Error Resume Next
    If Sock1.State <> sckClosed Then Sock1.Close
    AddText "You can't talk to " & ClientNick & " now because he/she is not connected to you."
    EnableWindow False
    RestartClient
    SStatus "Connection Closed"
    UpdateCaption
    If IsServerMode = True Then StartServer ServerPort
End Sub

Private Sub Sock1_Connect()
    On Error GoTo Er
    RestartClient
    'Sending of client info
    SendDataEx "SckGNk:", True 'Get other people's nickname
    Dim i As Long, j As Long, K As Long
    j = Val(GetSet("AutoEndDeadTimeOut", "3000"))
    K = Val(GetSet("AutoEndDeadWaitTime", "200"))
    Do Until ClientNick <> "" 'wait loop
        Wait K
        i = i + 1
        If i > j Then
            If GetSet("AutoEndDead", "1") = "1" Then
                End 'end if loop too long
            Else
                If MyMsgbox("The connection seems to be dead." & vbCrLf & vbCrLf & _
                                    "Do you want to end it now?", "DoNot8") = vbYes Then
                    End
                End If
            End If
        End If
    Loop
    EnableWindow True
    'Wait
    SendFont 0, True
    UpdateCaption
    AddText "You can start talking to " & ClientNick & " now."
    Exit Sub
Er:
    AddText "Connection Failed"
End Sub

Private Sub Sock1_ConnectionRequest(ByVal requestID As Long)
    On Error GoTo Er
    Dim i As Long, j As Long, K As Long
    Sock1.Close
    Sock1.Accept requestID
    EnableWindow True
    If Len(GetSet("ServerMsg")) > 0 Then _
        SendDataEx "SckMSG:Server Message: " & GetSet("ServerMsg") 'Send server message if set
    SendDataEx "SckGNk:", True 'get other people's nicknames
    j = Val(GetSet("AutoEndDeadTimeOut", "3000"))
    K = Val(GetSet("AutoEndDeadWaitTime", "200"))
    Do Until ClientNick <> "" 'wait loop
        Wait K
        i = i + 1
        If i > j Then
            If GetSet("AutoEndDead", "1") = "1" Then
                End 'end if loop too long
            Else
                If MyMsgbox("The connection seems to be dead." & vbCrLf & vbCrLf & _
                                    "Do you want to end it now?", "DoNot8") = vbYes Then
                    End
                End If
            End If
        End If
    Loop
    SendFont 0, True
    UpdateCaption
    AddText ClientNick & " has entered your chatroom."
    Exit Sub
Er:
    AddText "Error accepting connection for ID " & requestID
End Sub

Private Sub Sock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Incoming As String, Buffer As String
    Sock1.GetData Incoming 'Getting info
    If Len(ClientNick) > MaxNLen Then ClientNick = Mid$(ClientNick, 1, MaxNLen) 'Checking routine
    If Len(NickName) > MaxNLen Then NickName = Mid$(NickName, 1, MaxNLen) 'Checking routine
    If Len(Incoming) = 0 Then Exit Sub 'If it's nothing then don't bother doing the rest of the code
    Log "Incoming: " & Incoming
    EnableWindow
    If IsEncrypted(Incoming) = True Then Incoming = Encrypt(Incoming, False) 'Decryption of encrypted stuff
    Select Case Left$(Incoming, 7)
        'array for commands
        Case "MyName:"
            Buffer = StripTag(Incoming)
            ClientVer = GetString(Buffer, 1) 'used to be "incoming"
            If ClientNick <> "" And ClientNick <> GetString(Buffer) Then 'to prevent empty nicknames from popping up
                AddText ClientNick & " has changed his/her nickname to """ & GetString(Buffer) & """."
                Flash
            End If
            ClientNick = GetString(Buffer)
            lblContactInfo(0).Caption = ClientNick 'Added thingy!
            ClientBITS = Val(GetString(Buffer, 2))
            If Len(GetString(Buffer, 3)) > 0 Then 'This prevents a zero-image error
                If imgDP(0).Tag <> GetString(Buffer, 3) Then 'This stops the DP from being downloaded more than once
                    ChangeDP GetString(Buffer, 3), 1 'Download Client DP
                End If
            End If
            UpdateCaption
            If ClientBITS < BITSProtocol Then
                If MessageNotified(0) = False Then AddText "Warning: This contact might not be able to handle all commands that you send."
                MessageNotified(0) = True
            End If
        Case "SckGNk:"
            SendDataEx "MyName:" & NickName & "," & _
            App.ProductName & " V." & App.Major & "." & App.Minor & "." & App.Revision & "," & _
            BITSProtocol & "," & GetSet("DisplayPic", "http://www.kgv.net/blai/Images/manshead.bmp"), True
        Case "SckMSG:"
            AddText Incoming
            Flash
        Case "SckIMG:" 'Sending you an image
            AddImg StripTag(Incoming)
        Case "SckCMD:" 'cmd6 style lol
            'Call CMD6(GetString(StripTag(Incoming)))
            Call CMD6(StripTag(Incoming)) 'Why was I stripping that tag?
        Case "SckTyp:"
            If Len(ClientNick) = 0 Then SendDataEx "SckGNk:", True
            If StripTag(GetString(Incoming, 0)) = "1" Then
                SStatus ClientNick & " is typing a message."
            Else
                SStatus "Last message received at " & Now() & "."
            End If
        Case "MyFont:" 'new font transfer command
            SetFont Incoming
        Case "SckGft:"
            SendFont 0
        Case "SckErr:" 'If theres an error
            If GetSet("HideContactError", "1") = "0" Then 'If NOT HIDE errors (so it's.. SHOW xD)
                AddText ErrorProvider(GetString(StripTag(Incoming)), GetString(StripTag(Incoming), 1))
            End If
        Case Else
            If InStr(1, Incoming, "[!]") > 0 Then 'if theres a transfer error
                SendDataEx "SckErr:1", True
                Exit Sub
            End If
            If InStr(1, Incoming, "Sckft") = 0 Then
'                Incoming = DecodeEmoCodes(Incoming) 'Decode emo codes
'                AddEmos
                Buffer = Incoming

                If GetSet("AddName", "1") = "1" Then 'Add name
                    Dim ReplStr As String 'Dim for replacing nickname with style
                    ReplStr = Replace(GetSet("NickStyle", "[::]"), "::", ClientNick) 'This is the transformed style
                    Buffer = ReplStr & " " & Buffer  'Add name
                End If
                If GetSet("AddTime", "0") = "1" Then 'Add Time
                    Buffer = "(" & Format(Now, "Short Time") & ") " & Buffer 'Add Time
                End If
                If ClientNick = "" Then SendDataEx "SckGNk", True
                Dim OverFont As Integer
                OverFont = Val(GetSet("OverrideFont", "0")) 'this line adds the override font functionality
                Say Buffer, Abs(OverFont - 1)
                If GetSet("SpeakReceive", "0") = "1" Then
                    Voice.Speak Buffer, SVSFlagsAsync 'Speak!
                End If
                Flash True
                If Len(StrAutoMSG) > 0 Then 'auto message
                    If LastSendAutoMSG + 1 / 480 <= Now Then  'if 3 mins ago is later than last recorded time then
                        'SendData "[" & NickName & "] " & StrAutoMSG, False, (GetSet("EncryptData", "1") = 1)
                        SendData StrAutoMSG, False, (GetSet("EncryptData", "1") = 1)
                        Say "[" & NickName & " AutoMessage] " & StrAutoMSG
                        LastSendAutoMSG = Now
                    End If
                End If
            End If
    End Select
End Sub

Public Function Flash(Optional Now As Boolean)
    On Error Resume Next
    'If Me.WindowState = 1 Or Now = True Then FlashWindow Me.hwnd, 1
    If GetSet("FlashWindow", "1") = "0" Then Exit Function
    If GetActiveWindow <> Me.hwnd Or Now = True Then
        FlashWindow Me.hwnd, 1
    End If
End Function

Private Sub Sock2_ConnectionRequest(ByVal requestID As Long)
    On Error Resume Next
    Sock2.Close
    Log "Connection Accepted: " & requestID
    Sock2.Accept requestID
End Sub

Private Sub Sock2_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim Incoming As String
    Sock2.GetData Incoming 'Getting info
    If Len(Incoming) = 0 Then Exit Sub 'If it's nothing then don't bother doing the rest of the code
    Log "Incoming (S2): " & Incoming
    Select Case Left$(Incoming, 7)
        'array for commands
        Case "SckPng:"
            'if a client is not present the send the local port
            If Len(ClientNick) = 0 Then
                Sock2.SendData "SckPrt:" & Sock1.LocalPort
                Sock2.Close
            End If
    End Select
End Sub

Private Sub Sock3_Close()
    On Error Resume Next
    Sock3Connected = False
End Sub

Private Sub Sock3_Connect()
    On Error Resume Next
    Log "(S3) Connected"
    Sock3Connected = True
End Sub

Private Sub TimerRoutines_Timer()
    On Error Resume Next
    If GetActiveWindow <> Me.hwnd Then
            SendDataEx "SckTyp:0", True, False  'Reset other peope's Is Talking status
    End If
    SendFont 0, , False
    
    If GetSet("LogIP", "0") = "1" Then
        If Len(Sock1.RemoteHostIP) > 0 Then Log "IP: " & Sock1.RemoteHostIP
    End If
End Sub

Private Sub titAddContactEx_Click()
    On Error Resume Next
    frmAddContact.Show 1
    frmMain.LoadBuddies
End Sub

Private Sub titaddthisperson_Click()
    On Error Resume Next
    AddBuddy ClientNick, Sock1.RemoteHostIP
    frmMain.LoadBuddies
End Sub

Private Sub titAutoMSG_Click()
    On Error Resume Next
    Dim Buffer As String
    titAutoMSG.Checked = Not titAutoMSG.Checked
    If titAutoMSG.Checked = True Then
        Buffer = InputBox("Type a message that will be sent everytime when a client talks to you.", , _
                            GetSet("AutoMSG", "I'm away"))
        If Len(Buffer) > 0 Then
            StrAutoMSG = Buffer
            AddText "Auto message has been enabled."
        Else
            titAutoMSG.Checked = False
        End If
    Else
        StrAutoMSG = ""
        AddText "Auto message has been disabled."
    End If
End Sub

Private Sub titChangeDP_Click()
    On Error Resume Next
    Dim j As String
    j = InputBox("Enter the URL of the picture here:", , _
            GetSet("DisplayPic", "http://www.kgv.net/blai/Images/manshead.bmp"))
    If Len(j) = 0 Then Exit Sub
    ChangeDP j
End Sub

Public Function ChangeDP(URL As String, Optional YouOrMe As Integer = 0, Optional ForceUpdate As Boolean = False)
    On Error Resume Next
    If Len(URL) > 0 Then 'If there's nothing, whatever!
        If YouOrMe = 0 Then 'ME
            If ForceUpdate Or imgDP(1).Tag <> URL Then 'Forced update function patch
                frmChat.imgDP(1).Picture = LoadPicture(DownloadFile(URL, FindPath(App.Path, "dp.jpg")))
                SaveSet "DisplayPic", URL
                frmMain.imgDP(0).Picture = frmChat.imgDP(1).Picture
            End If
        ElseIf YouOrMe = 1 Then
            If ForceUpdate Or imgDP(0).Tag <> URL Then 'Forced update function patch
                imgDP(0).Picture = LoadPicture(DownloadFile(URL, FindPath(App.Path, "dp.jpg"))) 'Download Client DP
                frmMain.imgDP(1).Picture = imgDP(0).Picture 'The pic outside
                imgDP(0).Tag = URL
            End If
        End If
    End If
End Function

Private Sub titChgNN_Click()
    On Error Resume Next
    Dim Buffer As String, Buffer2 As String
    Buffer2 = NickName
    Buffer = InputBox("Change your Nickname to...", , NickName)
    If Len(Buffer) > 0 And Len(Buffer) <= MaxNLen Then
        Buffer = Replace(Buffer, ",", "")
        NickName = Buffer
        SendDataEx "SckNck:" & NickName, True
        AddText "You have changed your Nickname to """ & NickName & """."
        UpdateCaption
    Else
        If Len(Buffer) <> 0 Then
            MsgBox "Your nickname cannot be more than " & MaxNLen & " characters.", vbCritical
            titChgNN_Click
        End If
    End If
End Sub

Private Sub titClose_Click()
    On Error Resume Next
    Dim dx As Form
    For Each dx In Forms
        Unload dx
    Next
    End
End Sub

Private Sub titContactList_Click()
    On Error Resume Next
    frmMain.Handle_DblClick
End Sub

Private Sub titCopyCaption_Click()
    On Error Resume Next
    Clipboard.SetText lblContactInfo(Val(lblContactInfo(0).Tag)).Caption
End Sub

Private Sub titCopyIP_Click()
On Error Resume Next
    Clipboard.SetText Sock1.LocalIP
End Sub

Private Sub titPrefs_Click()
    On Error Resume Next
    frmPrefs.Show 1
End Sub

Private Sub titRefreshPic_Click()
    On Error Resume Next 'This is a sub for ME
    ChangeDP GetSet("DisplayPic"), , True
End Sub

Private Sub titRefreshPicEx_Click()
    On Error Resume Next 'This is a sub for contact
    ChangeDP imgDP(0).Tag, 1, True
End Sub

Private Sub titSaveLog_Click()
On Error GoTo Er
    Dim Buffer As String
    Buffer = SaveDialog("RTF Files (*.rtf)|*.rtf")
    If Len(Buffer) > 0 Then
        txtConvo.SaveFile Buffer, rtfRTF
        MsgBox "Log Saved", vbInformation
    End If
    Exit Sub
Er:
End Sub

Private Sub titSendAny_Click()
    On Error Resume Next
    Dim Buffer As String
    Buffer = InputBox("Send to contact:", , GetSet("LastSendAny"))
    If Len(Buffer) > 0 Then
        SendDataEx Buffer, True
        SaveSet "LastSendAny", IIf(IsEncrypted(Buffer), Encrypt(Buffer, False), Buffer)
    End If
End Sub

Private Sub titSendFile_Click()
    On Error GoTo Er
    cD1.Filter = "All Files (*.*)|*.*"
    cD1.ShowOpen
    If Len(cD1.FileName) > 0 Then
        SendDataEx "OpnTrf:", True
        Wait 200
        frmTransfer.Show
        Call frmTransfer.SendFile(cD1.FileName, Sock1.RemoteHostIP)
    End If
Er:
End Sub

Private Sub titSendImg_Click()
    On Error Resume Next
    Dim j As String
    j = InputBox("Enter URL here:")
    If Len(j) > 0 Then
        SendDataEx "SckIMG:" & j
        AddImg j
    End If
End Sub

Private Sub titTransparency_Click(Index As Integer)
On Error Resume Next
    If Index <> 10 Then
        MakeTransparent frmMain.hwnd, Index * 25.5
    Else
        MakeOpaque frmMain.hwnd 'Changed to frmMain cos...container form works.
    End If
End Sub

Private Sub txtConvo_GotFocus()
On Error Resume Next
    txtSendMsg.SetFocus
    If txtSendMsg.Locked = False Then 'Do, only if the control is not locked
        txtSendMsg.Text = txtSendMsg.Text & Chr$(KeyAscii)
        txtSendMsg.SelStart = Len(txtSendMsg.Text)
    End If
End Sub

Private Sub txtSendMsg_Change()
On Error Resume Next
    btnOK.Enabled = (Len(txtSendMsg.Text) > 0)
    If TypingModeSent = False Then
        SendDataEx "SckTyp:1"
        TypingModeSent = True
    ElseIf txtSendMsg.Text = "" Then
        SendDataEx "SckTyp:0"
        TypingModeSent = False
    End If
    If Len(StrAutoMSG) > 0 Then
        StrAutoMSG = ""
        titAutoMSG.Checked = False
        AddText "Auto message has been disabled."
    End If
End Sub

Public Function AddText(Default As String, Optional SystemNotification As Boolean = True)
    On Error Resume Next
    Dim SectionSrc As String
    If Left$(Default, 7) = "SckMSG:" Then Default = Mid$(Default, 8) 'Chop away the noobs
    With txtConvo
        LastLine = LastLine + 1 'Counting up of the last line variable
        If SystemNotification Then 'If this is a notification
            SectionSrc = "<font face=""arial"" size=""2"" color=""#7F7F7F""><i>" & Default & _
                                "</i></font><a name=""K" & LastLine & """> </a><br>"
        Else 'If this is NOT a notification
            Dim MyClr As String
            SectionSrc = "<font face=""arial"" size=""2"" color=""#" & GetHEXValue(0, 0, 0) & ">" & Default & _
                                "</font><a name=""K" & LastLine & """> </a><br>"
        End If
        WriteHTML SectionSrc
        txtConvo.Navigate2 FindPath(App.Path, "tmppage.html") & "#K" & LastLine
    End With
End Function

Public Function WriteHTML(WhatText As String)
    On Error Resume Next
    Dim FF As Integer
    FF = FreeFile
    Open FindPath(App.Path, "tmppage.html") For Append As #FF
        Print #FF, WhatText
    Close #FF
End Function

Public Function GetContactStatus(Index As Long)
On Error Resume Next
    Sock3.Close
    Sock3.Connect FetchItem(Index, 1), AvailabilityPort
    Log "GetStatus of " & Index
    While Sock3Connected <> True
        Wait 10
    Wend
    Sock3.SendData "GetSts:" & Index & "," & NickNamem
    Wait
    Sock3.Close
End Function

Private Sub txtSendMsg_Click()
    On Error Resume Next
    txtSendMsg_Change
End Sub
