VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4200
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   5250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3750
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      ScaleHeight     =   3750
      ScaleWidth      =   5250
      TabIndex        =   0
      Top             =   0
      Width           =   5250
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   624
         TabIndex        =   2
         Top             =   1728
         Width           =   3492
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         ForeColor       =   &H00FFFFFF&
         Height          =   372
         Left            =   -240
         TabIndex        =   1
         Top             =   1488
         Width           =   4380
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmServer.frm - Edit/Create a Server

Option Explicit

Private Sub Form_Load()
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  If App.Minor Mod 2 = 1 Then
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & "-Dev"
  Else
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  End If
  lblCopyright.Caption = "Copyright (C) 2001, The pgAdmin Development Team" & vbCrLf & "This software is released under the pgAdmin Public Licence"

End Sub

Private Sub Form_Resize()
  Me.Width = picLogo.Width
  Me.Height = picLogo.Height
  lblCopyright.Top = picLogo.Height - (lblCopyright.Height + 50)
  lblCopyright.Left = picLogo.Width - (lblCopyright.Width + 50)
  lblVersion.Left = picLogo.Width - (lblVersion.Width + 50)
  lblVersion.Top = picLogo.Height - (lblVersion.Height + lblCopyright.Height + 50)
End Sub
