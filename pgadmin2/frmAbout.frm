VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":08CA
   ScaleHeight     =   2835
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4050
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   2385
      Width           =   1230
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Copyright 2001, The pgAdmin Development Team"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   45
      TabIndex        =   4
      Top             =   1215
      Width           =   4275
   End
   Begin VB.Label lblLicence 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmAbout.frx":2D51
      ForeColor       =   &H0080FFFF&
      Height          =   720
      Left            =   45
      TabIndex        =   3
      Top             =   1575
      Width           =   5205
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   810
      Width           =   705
   End
   Begin VB.Label lblAppName 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AppName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   1350
      TabIndex        =   0
      Top             =   90
      Width           =   2610
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmAbout.frm - About Box.

Option Explicit

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmAbout.cmdOK_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmAbout.cmdOK_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmAbout.Form_Load()", etFullDebug

  lblAppName.Caption = App.Title
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & " Build " & App.Revision
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmAbout.Form_Load"
End Sub

