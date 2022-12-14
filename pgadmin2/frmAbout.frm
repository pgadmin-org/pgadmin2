VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3132
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4548
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3132
   ScaleWidth      =   4548
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2520
      Left            =   0
      Picture         =   "frmAbout.frx":0A02
      ScaleHeight     =   2520
      ScaleWidth      =   3648
      TabIndex        =   0
      Top             =   0
      Width           =   3648
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   135
         TabIndex        =   1
         Top             =   2310
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmAbout.frm - About Box.

Option Explicit

Private Sub picLogo_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmAbout.picLogo_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmAbout.picLogo_Click"
End Sub

Private Sub Form_Load()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmAbout.Form_Load()", etFullDebug

  lblVersion.Caption = ??TrasLang??("Version ") & App.Major & "." & App.Minor & "." & App.Revision
  If App.Minor Mod 2 = 1 Then
    lblVersion.Caption = ??TrasLang??("Version ") & App.Major & "." & App.Minor & "." & App.Revision & ??TrasLang??("-Dev")
  Else
    lblVersion.Caption = ??TrasLang??("Version ") & App.Major & "." & App.Minor & "." & App.Revision
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmAbout.Form_Load"
End Sub
