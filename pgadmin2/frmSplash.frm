VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3750
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
   ScaleHeight     =   3750
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
         ForeColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   135
         TabIndex        =   2
         Top             =   2745
         Width           =   930
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         ForeColor       =   &H00FFC0FF&
         Height          =   465
         Left            =   135
         TabIndex        =   1
         Top             =   3105
         Width           =   4965
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
  Me.Width = picLogo.Width
  Me.Height = picLogo.Height
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  If App.Minor Mod 2 = 1 Then
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & "-Dev"
  Else
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  End If
  lblCopyright.Caption = "Copyright (C) 2001, The pgAdmin Development Team" & vbCrLf & "This software is released under the pgAdmin Public Licence"
End Sub
