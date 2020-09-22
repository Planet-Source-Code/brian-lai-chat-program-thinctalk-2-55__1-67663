VERSION 5.00
Begin VB.Form frmTwip 
   BackColor       =   &H80000003&
   BorderStyle     =   0  '¨S¦³®Ø½u
   Caption         =   "Form1"
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   255
   LinkTopic       =   "Form1"
   ScaleHeight     =   255
   ScaleWidth      =   255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmTwip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
Me.Move MouseX * 15 + Me.Width + 15, MouseY * 15 + Me.Height + 15
End Sub
