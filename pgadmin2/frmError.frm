VERSION 5.00
Begin VB.Form frmError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error"
   ClientHeight    =   2475
   ClientLeft      =   4545
   ClientTop       =   3390
   ClientWidth     =   7110
   Icon            =   "frmError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   7110
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSendMail 
      Cancel          =   -1  'True
      Caption         =   "&Email Error"
      Height          =   375
      Left            =   45
      TabIndex        =   2
      Top             =   2025
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5940
      TabIndex        =   1
      Top             =   2025
      Width           =   1095
   End
   Begin VB.TextBox txterr 
      BackColor       =   &H8000000F&
      Height          =   1935
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmError.frm - Error form

' Note: No logging here

Option Explicit

Public Sub Initialise(lError As Long, szError As String, szRoutine As String)
Dim szTemp As String

  Me.Caption = App.Title & " Error"
  
  szTemp = "An error has occured in " & szRoutine & ":" & vbCrLf & vbCrLf
  szTemp = szTemp & "Number: " & lError & vbCrLf
  szTemp = szTemp & "Description: " & szError
  txterr.Text = szTemp
  
  'center form
  Me.Top = (Screen.Height - Me.Height) / 2
  Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdSendMail_Click()
Dim szMail As String
Dim szTemp As String

  szMail = "mailto:" & SUPPORT_EMAIL & "?subject=Error Message&body=" & txterr.Text
  szMail = Replace(szMail, " ", "%20")
  szMail = Replace(szMail, Chr(10), "%0" & Hex(10))
  szMail = Replace(szMail, Chr(13), "%0" & Hex(13))
  
  'open shell
  ShellExecute hwnd, "open", szMail, vbNullString, vbNullString, SW_SHOW
End Sub

