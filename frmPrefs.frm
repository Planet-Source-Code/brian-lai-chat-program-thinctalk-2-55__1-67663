VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrefs 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "ThincTalk - Preferences"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrefs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin MSComDlg.CommonDialog cD1 
      Left            =   120
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton btnUnloadMe 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   19
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   5520
      Width           =   975
   End
   Begin VB.ListBox LstTab 
      Height          =   5340
      IntegralHeight  =   0   'False
      ItemData        =   "frmPrefs.frx":000C
      Left            =   120
      List            =   "frmPrefs.frx":0022
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   5295
      Index           =   1
      Left            =   1560
      ScaleHeight     =   5295
      ScaleWidth      =   5175
      TabIndex        =   6
      Top             =   120
      Width           =   5175
      Begin VB.CheckBox chkInstReact 
         Caption         =   "Start when Windows Starts"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   65
         Top             =   4440
         Width           =   5175
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   1
         Left            =   4080
         TabIndex        =   52
         Text            =   "[::] "
         ToolTipText     =   "NickStyle,[::] "
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  '¸m¤¤¹ï»ô
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   51
         Text            =   "20"
         ToolTipText     =   "MaxNickLength,20"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   6
         Left            =   600
         TabIndex        =   36
         Text            =   "DP"
         ToolTipText     =   "DisplayPic,http://www.kgv.net/blai/Images/manshead.bmp"
         Top             =   960
         Width           =   3375
      End
      Begin VB.CommandButton btnTestDP 
         Caption         =   "Test"
         Height          =   375
         Left            =   4080
         TabIndex        =   35
         Top             =   930
         Width           =   975
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Remove extra line breaks, if any"
         Height          =   255
         Index           =   18
         Left            =   0
         TabIndex        =   22
         ToolTipText     =   "RemoveLineBreaks,1"
         Top             =   4080
         Width           =   5055
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   2
         Left            =   600
         MaxLength       =   20
         TabIndex        =   20
         Text            =   "Brian (LOLOL)"
         ToolTipText     =   "NickName,Guest"
         Top             =   360
         Width           =   3375
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Turn my words to leet (leet -> 1337)"
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "LeetData,0"
         Top             =   3720
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Log my contact's IP (per 10 seconds)"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   9
         ToolTipText     =   "LogIP,0"
         Top             =   3120
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Use Sounds if available"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   7
         ToolTipText     =   "UseSounds,1"
         Top             =   2760
         Width           =   5055
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "When a firewall/antivirus/whatever asks for permission, please click allow/permit/yes/whatever."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   390
         Index           =   19
         Left            =   240
         TabIndex        =   64
         Top             =   4680
         Width           =   4695
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Will be shown on the ""log"" tab page"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   61
         Top             =   3360
         Width           =   4650
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblURL 
         Alignment       =   2  '¸m¤¤¹ï»ô
         Caption         =   "Maximum length for Nicknames"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   56
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Preview: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   55
         Top             =   2400
         Width           =   675
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Style for showing nicknames: (::=Name)"
         Height          =   210
         Index           =   1
         Left            =   0
         TabIndex        =   54
         Top             =   2085
         Width           =   3315
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Limit to                    Characters"
         Height          =   210
         Index           =   0
         Left            =   600
         TabIndex        =   53
         Top             =   1485
         Width           =   2670
      End
      Begin VB.Image imgDP 
         Appearance      =   0  '¥­­±
         BorderStyle     =   1  '³æ½u©T©w
         Height          =   480
         Index           =   1
         Left            =   0
         Picture         =   "frmPrefs.frx":0061
         Stretch         =   -1  'True
         ToolTipText     =   "Contact's Display Picture"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Your Display Picture URL:"
         Height          =   210
         Index           =   7
         Left            =   600
         TabIndex        =   37
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Your Nickname:"
         Height          =   210
         Index           =   2
         Left            =   600
         TabIndex        =   21
         Top             =   120
         Width           =   1290
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   5295
      Index           =   0
      Left            =   1560
      ScaleHeight     =   5295
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton btnShellINI 
         Caption         =   "&Edit INI..."
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   4800
         Width           =   2175
      End
      Begin VB.ListBox LstCredits 
         Height          =   1530
         ItemData        =   "frmPrefs.frx":0B6B
         Left            =   120
         List            =   "frmPrefs.frx":0B84
         TabIndex        =   11
         Top             =   2640
         Width           =   4935
      End
      Begin VB.CommandButton btnWriteXPVS 
         Caption         =   "Use XP Visual Styles"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "List of beta testers: (thank you so much!)"
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   3465
      End
      Begin VB.Image Image2 
         Height          =   645
         Left            =   120
         Picture         =   "frmPrefs.frx":0BD7
         Top             =   120
         Width           =   2190
      End
      Begin VB.Label lblProdDes 
         BackStyle       =   0  '³z©ú
         Caption         =   "Description"
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   4935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   5040
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblProdVer 
         BackStyle       =   0  '³z©ú
         Caption         =   "Version "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4935
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   5295
      Index           =   3
      Left            =   1560
      ScaleHeight     =   5295
      ScaleWidth      =   5175
      TabIndex        =   23
      Top             =   120
      Width           =   5175
      Begin VB.CheckBox chkOpt 
         Caption         =   "Speak what I send"
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   30
         ToolTipText     =   "SpeakSend,0"
         Top             =   720
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Speak what the contact types"
         Height          =   255
         Index           =   13
         Left            =   0
         TabIndex        =   29
         ToolTipText     =   "SpeakReceive,0"
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "ThincTalk can use a default speech engine (""Microsoft Sam"") to read out whatever is sent or received."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   390
         Index           =   17
         Left            =   240
         TabIndex        =   62
         Top             =   120
         Width           =   4650
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   5295
      Index           =   2
      Left            =   1560
      ScaleHeight     =   5295
      ScaleWidth      =   5175
      TabIndex        =   8
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton btnFindFile 
         Caption         =   "Browse"
         Height          =   375
         Index           =   5
         Left            =   4200
         TabIndex        =   69
         Top             =   330
         Width           =   975
      End
      Begin VB.CommandButton btnFindFile 
         Caption         =   "Browse"
         Height          =   375
         Index           =   7
         Left            =   4200
         TabIndex        =   68
         Top             =   1050
         Width           =   975
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   7
         Left            =   0
         TabIndex        =   66
         Text            =   "SkinFile"
         ToolTipText     =   "SkinFile,"
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   5
         Left            =   0
         TabIndex        =   33
         Text            =   "BkgImg"
         ToolTipText     =   "BackgroundImg,http://www.kgv.net/blai/Images/Bkg.bmp"
         Top             =   360
         Width           =   4095
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Timestamp messages"
         Height          =   255
         Index           =   20
         Left            =   0
         TabIndex        =   32
         ToolTipText     =   "AddTime,0"
         Top             =   2400
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Force conversation window to be enabled"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   31
         ToolTipText     =   "ForceEnableWindow,0"
         Top             =   4440
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Conversation Window: black border"
         Height          =   255
         Index           =   14
         Left            =   0
         TabIndex        =   18
         ToolTipText     =   "ConvoBorder,1"
         Top             =   4080
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Display sender's name on messages"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   17
         ToolTipText     =   "AddName,1"
         Top             =   2040
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Flash Window when someone talks to you"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   16
         ToolTipText     =   "FlashWindow,1"
         Top             =   2760
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Always Enable Send Button"
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   15
         ToolTipText     =   "ForceSend,0"
         Top             =   3120
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Override contact's font"
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   14
         ToolTipText     =   "OverrideFont,0"
         Top             =   3480
         Width           =   5055
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Skin File... (*.ini)"
         Height          =   210
         Index           =   20
         Left            =   0
         TabIndex        =   67
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "The contact's font will look just like yours."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   18
         Left            =   240
         TabIndex        =   63
         Top             =   3720
         Width           =   4650
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Convo Window Background Image:"
         Height          =   210
         Index           =   6
         Left            =   0
         TabIndex        =   34
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   5295
      Index           =   4
      Left            =   1560
      ScaleHeight     =   5295
      ScaleWidth      =   5175
      TabIndex        =   24
      Top             =   120
      Width           =   5175
      Begin VB.CheckBox chkOpt 
         Caption         =   "Suppress all system notices"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   28
         ToolTipText     =   "SuppressSysNotice,0"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "When an interactive command is received, show it"
         Height          =   255
         Index           =   15
         Left            =   0
         TabIndex        =   27
         ToolTipText     =   "ShowCMD6ToClient,0"
         Top             =   720
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Hide contact's errors"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   26
         ToolTipText     =   "HideContactError,1"
         Top             =   1920
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Show Cursor Position in Remote Desktop"
         Height          =   255
         Index           =   19
         Left            =   0
         TabIndex        =   25
         ToolTipText     =   "OpenTwip,0"
         Top             =   120
         Width           =   5055
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Hides the error messages sent from the client"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   15
         Left            =   480
         TabIndex        =   60
         Top             =   2160
         Width           =   4650
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Gray system text will not be shown even when there's an error"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   59
         Top             =   1560
         Width           =   4650
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Will be shown on the conversation area"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   58
         Top             =   960
         Width           =   4650
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "A blue square will be positioned somewhere next to the cursor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   57
         Top             =   360
         Width           =   4650
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   5295
      Index           =   5
      Left            =   1560
      ScaleHeight     =   5295
      ScaleWidth      =   5175
      TabIndex        =   38
      Top             =   120
      Width           =   5175
      Begin VB.CheckBox chkOpt 
         Caption         =   "Do not connect to erroneous ports (e.g. 98765)"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   44
         ToolTipText     =   "BlockErroneousPorts,1"
         Top             =   120
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Send information even if I am connected to nobody"
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   43
         ToolTipText     =   "ForceSendData,0"
         Top             =   720
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Encrypt Sent Data"
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   42
         ToolTipText     =   "EncryptData,1"
         Top             =   1320
         Width           =   5055
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  '¾a¥k¹ï»ô
         Height          =   315
         Index           =   4
         Left            =   3000
         MaxLength       =   20
         TabIndex        =   41
         Text            =   "100"
         ToolTipText     =   "AutoEndDeadWaitTime,100"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  '¾a¥k¹ï»ô
         Height          =   315
         Index           =   3
         Left            =   3000
         MaxLength       =   20
         TabIndex        =   40
         Text            =   "4000"
         ToolTipText     =   "AutoEndDeadTimeOut,4000"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Auto-End dead connections if idle"
         Height          =   255
         Index           =   17
         Left            =   0
         TabIndex        =   39
         ToolTipText     =   "AutoEndDead,1"
         Top             =   1920
         Width           =   5055
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Advanced info: you are telling Winsock to send to nowhere"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   50
         Top             =   960
         Width           =   4650
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Hello -> Igolp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   49
         Top             =   1560
         Width           =   4650
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Auto-end stops your CPU from burning."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   48
         Top             =   2160
         Width           =   4650
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "A correct port is a number between 1 and 65535."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   4650
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Wait time for each interval (ms): "
         Height          =   210
         Index           =   4
         Left            =   180
         TabIndex        =   46
         Top             =   2805
         Width           =   2730
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Wait for (Times): "
         Height          =   210
         Index           =   3
         Left            =   180
         TabIndex        =   45
         Top             =   2445
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btnFindFile_Click(Index As Integer)
    On Error GoTo OnOz
    cD1.Filter = "Skin Files (*.ini)|*.ini"
    cD1.ShowOpen
    txtData(Index).Text = cD1.FileName
OnOz:
End Sub

Private Sub btnOK_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To chkOpt.UBound Step 1 'Save Settings
        SaveSet GetString(chkOpt(i).ToolTipText), Str$(chkOpt(i).Value)
    Next
    For i = 0 To txtData.UBound Step 1 'Save Settings
        SaveSet GetString(txtData(i).ToolTipText), txtData(i).Text
    Next
    'load my dp
    frmChat.ChangeDP txtData(6).Text
    Unload Me
End Sub

Private Sub btnShellINI_Click()
    On Error Resume Next
    Shell "notepad " & SettingsFile, vbNormalFocus
End Sub

Private Sub btnTestDP_Click()
    On Error Resume Next
    imgDP(1).Picture = LoadPicture(DownloadFile(txtData(6).Text, FindPath(App.Path, "dp.jpg")))
End Sub

Private Sub btnUnloadMe_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub btnWriteXPVS_Click()
    On Error Resume Next
    If MsgBox("This function will write the manifest file again to show the Windows XP Visual Styles if applicable.", _
    vbYesNo + vbQuestion) = vbNo Then Exit Sub
    XPVB
    MsgBox "Manifest Written. Please restart " & App.ProductName & ".", vbInformation
End Sub

Private Sub chkInstReact_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0 'The Start when windows starts thing
            If CheStart.Value = 1 Then
                SaveRegString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", App.ProductName, FindPath(App.Path, App.ProductName & ".exe")
            Else
                DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", App.ProductName
            End If
    End Select
End Sub

Private Sub chkOpt_Click(Index As Integer)
    On Error Resume Next
    chkOpt(16).Enabled = (chkOpt(3).Value = 0)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    'lblProdName.Caption = App.ProductName
    lblProdVer.Caption = App.ProductName & " " & MyVer
    lblProdDes.Caption = App.ProductName & " " & MyVer & ", all rights reserved by Thinc." & vbCrLf & _
    "Made by Brian Lai" & vbCrLf & "http://thinc.no-ip.info" & vbCrLf & _
    "Help from: 1997 SoftCircuits Programming"
    For i = 0 To chkOpt.UBound Step 1 'Load Settings
        chkOpt(i).Value = GetSet(GetString(chkOpt(i).ToolTipText), GetString(chkOpt(i).ToolTipText, 1))
    Next
    For i = 0 To txtData.UBound Step 1 'Load Settings
        txtData(i).Text = GetSet(GetString(txtData(i).ToolTipText), GetString(txtData(i).ToolTipText, 1))
    Next
    SkinForm Me
    'Load for some checkbox...
    Dim StartUp As String
    StartUp = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", App.ProductName)
    chkInstReact(0).Value = IIf(StartUp = vbNullString, 0, 1)
End Sub

Private Sub lblURL_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            GoToTab 2
    End Select
End Sub

Private Sub LstTab_Click()
    On Error Resume Next
    picTabSwitch(LstTab.ListIndex).ZOrder 0
End Sub

Public Function GoToTab(Index As Integer)
    On Error Resume Next
    picTabSwitch(Index).ZOrder 0
End Function

Private Sub txtData_Change(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            txtData(2).MaxLength = Val(txtData(0).Text)
            txtData(2).Text = Left$(txtData(2).Text, Val(txtData(0).Text))
        Case 1
            lblNote(12).Caption = "Preview: " & Replace(txtData(Index).Text, "::", frmChat.NickName) & "Hello"
    End Select
End Sub
