VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWizard 
   Appearance      =   0  '¥­­±
   BackColor       =   &H00E0E0E0&
   Caption         =   "Connection Wizard"
   ClientHeight    =   5475
   ClientLeft      =   5595
   ClientTop       =   3240
   ClientWidth     =   7185
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
   Icon            =   "frmGuide.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   7185
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.PictureBox PicMovable 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   960
      ScaleHeight     =   4065
      ScaleWidth      =   5265
      TabIndex        =   0
      Top             =   720
      Width           =   5295
      Begin VB.CommandButton btnCancel 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton btnBackNForth 
         Caption         =   "&Next"
         Default         =   -1  'True
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   2
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton btnBackNForth 
         Caption         =   "&Back"
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   1
         Top             =   3600
         Width           =   975
      End
      Begin MSWinsockLib.Winsock Sock1 
         Left            =   1200
         Top             =   3600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   3485
         Index           =   0
         Left            =   0
         ScaleHeight     =   3480
         ScaleWidth      =   5295
         TabIndex        =   4
         Top             =   0
         Width           =   5295
         Begin VB.TextBox txtNickName 
            Alignment       =   2  '¸m¤¤¹ï»ô
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
            Left            =   1800
            TabIndex        =   5
            Text            =   "Guest"
            Top             =   1485
            Width           =   1935
         End
         Begin VB.Label lblNotation 
            AutoSize        =   -1  'True
            BackStyle       =   0  '³z©ú
            Caption         =   "Your Nickname:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   1530
            Width           =   1290
         End
         Begin VB.Label lblNotation 
            BackStyle       =   0  '³z©ú
            Caption         =   "Start talking to your friends."
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
            Left            =   360
            TabIndex        =   7
            Top             =   1200
            Width           =   3495
         End
         Begin VB.Image imgLogo 
            Height          =   645
            Left            =   1440
            Picture         =   "frmGuide.frx":1982
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2190
         End
         Begin VB.Label lblNotation 
            AutoSize        =   -1  'True
            BackStyle       =   0  '³z©ú
            Caption         =   "You cannot have commas (,) in your nickname. All commas will be removed."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   420
            Index           =   10
            Left            =   360
            TabIndex        =   6
            Top             =   2280
            Width           =   3960
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  '¨S¦³®Ø½u
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3485
         Index           =   3
         Left            =   0
         ScaleHeight     =   3480
         ScaleWidth      =   5295
         TabIndex        =   28
         Top             =   0
         Width           =   5295
         Begin VB.Frame Frame2 
            Caption         =   "Required information"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   5055
            Begin VB.PictureBox Picture3 
               BorderStyle     =   0  '¨S¦³®Ø½u
               Height          =   1455
               Left            =   120
               ScaleHeight     =   1455
               ScaleWidth      =   4815
               TabIndex        =   32
               Top             =   240
               Width           =   4815
               Begin VB.CommandButton btnSearch 
                  Caption         =   "Find Network..."
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   0
                  TabIndex        =   54
                  Top             =   1080
                  Width           =   1575
               End
               Begin VB.TextBox txtClientPort 
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
                  Left            =   3960
                  TabIndex        =   34
                  Text            =   "8918"
                  Top             =   240
                  Width           =   855
               End
               Begin VB.TextBox txtClientIP 
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
                  Left            =   0
                  TabIndex        =   33
                  Text            =   "127.0.0.1"
                  Top             =   240
                  Width           =   3375
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '³z©ú
                  Caption         =   "Port:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Index           =   1
                  Left            =   3480
                  TabIndex        =   36
                  Top             =   277
                  Width           =   405
               End
               Begin VB.Label Label2 
                  BackStyle       =   0  '³z©ú
                  Caption         =   "Connect to server: "
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
                  Left            =   0
                  TabIndex        =   35
                  Top             =   0
                  Width           =   3615
               End
            End
         End
         Begin VB.TextBox txtContactName 
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
            Height          =   285
            Left            =   2640
            MaxLength       =   50
            TabIndex        =   29
            Text            =   "My Friend"
            Top             =   3000
            Width           =   2535
         End
         Begin VB.CheckBox chkAddContact 
            Caption         =   "Add as contact: "
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
            Left            =   120
            TabIndex        =   30
            Top             =   3000
            Width           =   3975
         End
         Begin VB.Label lblNotation 
            AutoSize        =   -1  'True
            BackStyle       =   0  '³z©ú
            Caption         =   "Fill in the required information to start chatting."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   5055
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   3485
         Index           =   4
         Left            =   0
         ScaleHeight     =   3480
         ScaleWidth      =   5295
         TabIndex        =   20
         Top             =   0
         Width           =   5295
         Begin VB.Frame Frame2 
            Caption         =   "Talk to buddies"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   5055
            Begin VB.ListBox lstContacts 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2460
               IntegralHeight  =   0   'False
               ItemData        =   "frmGuide.frx":21B2
               Left            =   120
               List            =   "frmGuide.frx":21B4
               TabIndex        =   23
               Top             =   240
               Width           =   3135
            End
            Begin VB.ListBox lstBuddyIndex 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1740
               ItemData        =   "frmGuide.frx":21B6
               Left            =   2160
               List            =   "frmGuide.frx":21B8
               TabIndex        =   22
               Top             =   360
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label lblClientInfo 
               BackStyle       =   0  '³z©ú
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   0
               Left            =   3360
               TabIndex        =   26
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label lblClientInfo 
               BackStyle       =   0  '³z©ú
               Height          =   495
               Index           =   1
               Left            =   3360
               TabIndex        =   25
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label lblClientInfo 
               BackStyle       =   0  '³z©ú
               Height          =   495
               Index           =   2
               Left            =   3360
               TabIndex        =   24
               Top             =   1200
               Width           =   1575
            End
         End
         Begin VB.Label lblNotation 
            AutoSize        =   -1  'True
            BackStyle       =   0  '³z©ú
            Caption         =   "Double Click on a contact to talk to him/her."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   5055
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   3485
         Index           =   2
         Left            =   0
         ScaleHeight     =   3480
         ScaleWidth      =   5295
         TabIndex        =   40
         Top             =   0
         Width           =   5295
         Begin VB.Frame Frame3 
            Caption         =   "Server Options"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   5055
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  '¨S¦³®Ø½u
               Height          =   2415
               Left            =   120
               ScaleHeight     =   2415
               ScaleWidth      =   4815
               TabIndex        =   43
               Top             =   240
               Width           =   4815
               Begin VB.TextBox txtServCustomPort 
                  Alignment       =   2  '¸m¤¤¹ï»ô
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3120
                  TabIndex        =   52
                  Text            =   "8918"
                  Top             =   1200
                  Width           =   1575
               End
               Begin VB.TextBox txtServerMsg 
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
                  Left            =   120
                  TabIndex        =   51
                  Text            =   "Hi"
                  Top             =   2040
                  Width           =   4575
               End
               Begin VB.OptionButton OptServer 
                  Caption         =   "Listen to other port:"
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
                  TabIndex        =   49
                  Top             =   1200
                  Width           =   4575
               End
               Begin VB.OptionButton OptServer 
                  Caption         =   "Listen to default port (Recommended)"
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
                  TabIndex        =   48
                  Top             =   840
                  Value           =   -1  'True
                  Width           =   4575
               End
               Begin VB.CommandButton btnCheckIP 
                  Caption         =   "Check"
                  Height          =   375
                  Left            =   3840
                  TabIndex        =   47
                  Top             =   315
                  Width           =   855
               End
               Begin VB.TextBox txtYourIP 
                  Alignment       =   2  '¸m¤¤¹ï»ô
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
                  Left            =   1080
                  TabIndex        =   46
                  Text            =   "Text1"
                  Top             =   360
                  Width           =   2655
               End
               Begin VB.Label lblNotation 
                  BackStyle       =   0  '³z©ú
                  Caption         =   "Server Message: (Optional)"
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
                  Index           =   12
                  Left            =   120
                  TabIndex        =   50
                  Top             =   1800
                  Width           =   3495
               End
               Begin VB.Label lblNotation 
                  BackStyle       =   0  '³z©ú
                  Caption         =   "Your IP is"
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
                  Index           =   11
                  Left            =   120
                  TabIndex        =   45
                  Top             =   375
                  Width           =   3495
               End
               Begin VB.Label lblNotation 
                  BackStyle       =   0  '³z©ú
                  Caption         =   "Note: your friends will need to know this to connect to you."
                  ForeColor       =   &H00404040&
                  Height          =   255
                  Index           =   9
                  Left            =   120
                  TabIndex        =   44
                  Top             =   0
                  Width           =   4695
               End
            End
         End
         Begin VB.Label lblNotation 
            BackStyle       =   0  '³z©ú
            Caption         =   "You have chosen to set up a server."
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
            Index           =   6
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   3485
         Index           =   1
         Left            =   0
         ScaleHeight     =   3480
         ScaleWidth      =   5295
         TabIndex        =   9
         Top             =   0
         Width           =   5295
         Begin VB.CheckBox chkRememberTypeChoice 
            Caption         =   "Remember my choice"
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
            Left            =   240
            TabIndex        =   53
            Top             =   3120
            Width           =   4815
         End
         Begin VB.Frame Frame1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   5055
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  '¨S¦³®Ø½u
               Height          =   2055
               Left            =   120
               ScaleHeight     =   2055
               ScaleWidth      =   4815
               TabIndex        =   12
               Top             =   240
               Width           =   4815
               Begin VB.OptionButton OptType 
                  Caption         =   "Talk to my saved contacts"
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
                  Left            =   0
                  TabIndex        =   15
                  Top             =   720
                  Value           =   -1  'True
                  Width           =   4215
               End
               Begin VB.OptionButton OptType 
                  Caption         =   "Talk to someone not on my contact list"
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
                  Left            =   0
                  TabIndex        =   14
                  Top             =   1320
                  Width           =   4215
               End
               Begin VB.OptionButton OptType 
                  Caption         =   "Be a Server"
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
                  Left            =   0
                  TabIndex        =   13
                  Top             =   120
                  Width           =   4215
               End
               Begin VB.Label lblNotation 
                  BackStyle       =   0  '³z©ú
                  Caption         =   "Make a chatroom and let others join."
                  ForeColor       =   &H00404040&
                  Height          =   255
                  Index           =   3
                  Left            =   360
                  TabIndex        =   18
                  Top             =   360
                  Width           =   4335
               End
               Begin VB.Label lblNotation 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '³z©ú
                  Caption         =   "Choose from a list of contacts that you previously saved."
                  ForeColor       =   &H00404040&
                  Height          =   390
                  Index           =   4
                  Left            =   360
                  TabIndex        =   17
                  Top             =   960
                  Width           =   4410
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblNotation 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '³z©ú
                  Caption         =   "You can talk to anyone by entering his/her computer IP address."
                  ForeColor       =   &H00404040&
                  Height          =   390
                  Index           =   5
                  Left            =   360
                  TabIndex        =   16
                  Top             =   1560
                  Width           =   4410
                  WordWrap        =   -1  'True
               End
            End
         End
         Begin VB.CommandButton btnTalkToMe 
            Caption         =   "Talk to Brian"
            Height          =   435
            Left            =   3720
            TabIndex        =   10
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label lblNotation 
            AutoSize        =   -1  'True
            BackStyle       =   0  '³z©ú
            Caption         =   "What would you like to do?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   2280
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Tag             =   "Y:3480"
         X1              =   0
         X2              =   5280
         Y1              =   3480
         Y2              =   3480
      End
   End
   Begin VB.Label lblVer 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   39
      Top             =   5160
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   1
      Left            =   5160
      Picture         =   "frmGuide.frx":21BA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   0
      Left            =   0
      Picture         =   "frmGuide.frx":3720
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label lblWeb 
      Alignment       =   1  '¾a¥k¹ï»ô
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Visit Website"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6000
      TabIndex        =   38
      ToolTipText     =   "http://thinc.no-ip.info"
      Top             =   5160
      Width           =   1065
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClickOption As Integer
Dim TopMostTab As Integer

Private Sub btnBackNForth_Click(Index As Integer)
    On Error Resume Next
    Dim i As Integer
    Dim Buffer1 As String, Buffer2 As String
    'Universal Event handlers
    If btnBackNForth(Index).Caption = "Close" Then End
    '/Universal Event handlers
    If TopMostTab + Index - 2 > picTab.UBound Or TopMostTab + Index - 2 < picTab.LBound Then Exit Sub 'Tab Controllers
    TopMostTab = TopMostTab + Index - 2
    picTab(TopMostTab).ZOrder 0
    If TopMostTab = 1 Then 'Per-Tab Event handlers
        If Len(txtNickName.Text) < 1 Then
            MsgBox "You might enter a nickname.", vbCritical
            btnBackNForth_Click 1
        End If
    ElseIf TopMostTab = 2 Then 'This is the client-server selection

        For i = 0 To 2 Step 1
            If OptType(i).Value = True Then
                picTab(i + 2).ZOrder 0
                Exit For
            End If
        Next
    ElseIf TopMostTab > 2 Then 'if everything is done
        btnBackNForth(3).Enabled = False
        SaveSet "Nickname", txtNickName.Text
        SaveSet "ServerMsg", txtServerMsg.Text
        txtNickName.Text = Replace(txtNickName.Text, ",", "")
        frmChat.NickName = txtNickName.Text
        Select Case ClickOption
            Case 0
                If txtClientPort.Text = "" Then txtClientPort.Text = DefaultPort
                Call frmChat.StartServer(txtServCustomPort.Text)
            Case 1
                If chkAddContact.Value = 1 Then 'add contact
                    AddBuddy txtContactName.Text, txtClientIP.Text, txtClientPort.Text
                End If
                Call frmChat.StartClient(txtClientIP.Text, Val(txtClientPort.Text))
            Case 2
                If lstContacts.ListIndex < 0 Then lstContacts.ListIndex = 0
                Buffer1 = FetchItem(Val(lstBuddyIndex.List(lstContacts.ListIndex)), 1)
                Buffer2 = FetchItem(Val(lstBuddyIndex.List(lstContacts.ListIndex)), 2)
                Call frmChat.StartClient(Buffer1, Val(Buffer2)) 'starts new client
        End Select
        frmChat.Show
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
On Error Resume Next
Dim dx As Form
For Each dx In Forms
    Unload dx
Next
End
End Sub

Private Sub btnCheckIP_Click()
    On Error Resume Next
    If Left(txtYourIP.Text, 7) = "192.168" Then  'warning
        MsgBox "Warning: the displayed Internet adress (""IP"") is internal!" & vbCrLf & _
                    "Others from the Internet will not able to connect to you using this IP.", _
                    vbExclamation
    ElseIf txtYourIP.Text = "127.0.0.1" Then
        MsgBox "Warning: the displayed Internet adress (""IP"") is a loopback address!" & vbCrLf & _
                    "Others will not able to connect to you using this IP.", _
                    vbExclamation
    Else
        MsgBox "Your IP seems to be good!", vbInformation
    End If
End Sub

Private Sub btnTalkToMe_Click()
    On Error Resume Next
    OptType(1).Value = True
    btnBackNForth_Click 3
    txtClientIP.Text = GetSet("TalkToDefaultIP", "thinc.myvnc.com", "Global")
    txtClientPort.Text = GetSet("TalkToDefaultPort", "8918", "Global")
    btnBackNForth_Click 3
End Sub

Private Sub chkAddContact_Click()
    On Error Resume Next
    txtContactName.Enabled = (chkAddContact.Value = 1)
End Sub

Private Sub chkRememberTypeChoice_Click()
    On Error Resume Next
    SaveSet "RememberTypeChoice", Val(chkRememberTypeChoice.Value)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim i As Long, j As Integer
    Dim Buffer As String, K As String
    frmMain.picSideBar.Visible = False
    'Me.BackColor = RGB(78, 113, 167)
    Me.BackColor = RGB(108, 108, 108)
    Sock1Ready = True 'do not remove - otherwise it wont start sending info.
    ClickOption = 2
    Frame1.Caption = "Set up Chatroom: [" & Sock1.LocalIP & "]"
    txtYourIP.Text = Sock1.LocalIP
    txtNickName.Text = GetSet("Nickname", UserName)
    txtServerMsg.Text = GetSet("ServerMsg")
    chkRememberTypeChoice.Value = Val(GetSet("RememberTypeChoice", "1"))
    If chkRememberTypeChoice.Value = 1 Then
        OptType(Val(GetSet("ServerTypeChoice", "2"))).Value = True ' if remember settings, load it
    End If
    btnTalkToMe.Caption = "Talk to " & GetSet("TalkToDefault", "Brian", "Global")
    K = GetSet("TTalkLogo", , "Global")
    If Len(K) > 0 Then imgLogo.Picture = LoadPicture(K) 'load the "company logo" (customizable...)
    lstContacts.Clear
    lstBuddyIndex.Clear
    For i = 0 To MaxBuddies
        Buffer = FetchItem(i, 0)
        If Len(Buffer) > 0 Then
            lstContacts.AddItem Buffer
            lstBuddyIndex.AddItem i
        End If
    Next
    lblVer.Caption = MyVer
    SkinForm Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.PicMovable.Move (Me.ScaleWidth - PicMovable.Width) / 2, (Me.ScaleHeight - PicMovable.Height) / 2
    lblWeb.Move Me.ScaleWidth - lblWeb.Width - 120, Me.ScaleHeight - lblWeb.Height - 120
    lblVer.Move 120, Me.ScaleHeight - lblWeb.Height - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmMain.picSideBar.Visible = True
End Sub


Private Sub lblWeb_Click()
    On Error Resume Next
    Shell "Explorer http://thinc.no-ip.info", vbNormalFocus
End Sub

Private Sub lstContacts_Click()
    On Error Resume Next
    lblClientInfo(0).Caption = lstContacts.List(lstContacts.ListIndex)
    lblClientInfo(1).Caption = "IP: " & FetchItem(Val(lstBuddyIndex.List(lstContacts.ListIndex)), 1)
    lblClientInfo(2).Caption = "Port " & FetchItem(Val(lstBuddyIndex.List(lstContacts.ListIndex)), 2)
End Sub

Private Sub lstContacts_DblClick()
    On Error Resume Next
    Dim Buffer1 As String, Buffer2 As String
    If Len(txtNickName.Text) < 1 Then
        MsgBox "You might enter a nickname.", vbCritical
        Exit Sub
    Else
        SaveSet "Nickname", txtNickName.Text
    End If
    frmChat.NickName = txtNickName.Text
    Buffer1 = FetchItem(Val(lstBuddyIndex.List(lstContacts.ListIndex)), 1)
    Buffer2 = FetchItem(Val(lstBuddyIndex.List(lstContacts.ListIndex)), 2)
    Call frmChat.StartClient(Buffer1, Val(Buffer2)) 'starts new client
    frmChat.Show
    Unload Me
End Sub

Private Sub OptServer_Click(Index As Integer)
    On Error Resume Next
    txtServCustomPort.Enabled = (OptServer(1).Value = True)
    If OptServer(0).Value = True Then txtServCustomPort.Text = DefaultPort
End Sub

Private Sub OptType_Click(Index As Integer)
    On Error Resume Next
    Frame2(Index).ZOrder 0
    ClickOption = Index
    If chkRememberTypeChoice.Value = 1 Then 'if record value
        SaveSet "ServerTypeChoice", Str(Index)
    End If
End Sub

Private Sub txtServCustomPort_LostFocus()
On Error Resume Next
If Val(txtServCustomPort.Text) <= 0 Or Val(txtServCustomPort.Text) > 65535 Then
    MsgBox "The port " & Val(txtServCustomPort.Text) & " is not a valid port. Please try again.", vbCritical
    txtServCustomPort.Text = DefaultPort
End If
End Sub

