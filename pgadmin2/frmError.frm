VERSION 5.00
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error"
   ClientHeight    =   5790
   ClientLeft      =   4440
   ClientTop       =   4215
   ClientWidth     =   7425
   Icon            =   "frmError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7425
   Begin HighlightBox.TBX txtMore 
      Height          =   2535
      Left            =   1305
      TabIndex        =   10
      Top             =   3195
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   4471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
   End
   Begin HighlightBox.TBX txtErr 
      Height          =   1770
      Left            =   45
      TabIndex        =   3
      Top             =   945
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   3122
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
   End
   Begin VB.PictureBox PictList 
      BackColor       =   &H00808080&
      Height          =   2535
      Left            =   45
      ScaleHeight     =   2475
      ScaleWidth      =   1185
      TabIndex        =   13
      Top             =   3210
      Width           =   1245
      Begin VB.CommandButton cmdInfo 
         Caption         =   "&Error Info."
         Height          =   375
         Index           =   4
         Left            =   45
         TabIndex        =   9
         ToolTipText     =   "Display error information."
         Top             =   2025
         Width           =   1095
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "D&river Info."
         Height          =   375
         Index           =   3
         Left            =   45
         TabIndex        =   8
         ToolTipText     =   "Display ODBC driver information."
         Top             =   1530
         Width           =   1095
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "&Db Info."
         Height          =   375
         Index           =   2
         Left            =   45
         TabIndex        =   7
         ToolTipText     =   "Display database information."
         Top             =   1035
         Width           =   1095
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "&App Info."
         Height          =   375
         Index           =   1
         Left            =   45
         TabIndex        =   6
         ToolTipText     =   "Display application information."
         Top             =   540
         Width           =   1095
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "&System Info."
         Height          =   375
         Index           =   0
         Left            =   45
         TabIndex        =   5
         ToolTipText     =   "Display system information."
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "D&etails >>"
      Height          =   375
      Left            =   90
      TabIndex        =   4
      ToolTipText     =   "Show extended details of the error."
      Top             =   2790
      Width           =   1095
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      ToolTipText     =   "Ignore the error and continue running pgAdmin."
      Top             =   2775
      Width           =   1095
   End
   Begin VB.CommandButton cmdSendMail 
      Caption         =   "Send &Mail"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      ToolTipText     =   "Send details of this error to the support mailing list."
      Top             =   2775
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      ToolTipText     =   "Exit from pgAdmin."
      Top             =   2775
      Width           =   1095
   End
   Begin VB.Image imgErr 
      Height          =   480
      Left            =   6720
      Picture         =   "frmError.frx":0442
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Error Handler"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "An Error Occured in pgAdmin2;"
      Height          =   195
      Left            =   945
      TabIndex        =   11
      Top             =   480
      Width           =   2190
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   45
      Top             =   45
      Width           =   7335
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

Private Const MailSupport = "pgadmin-support@postgresql.org"

Dim objError As clsError

Public Sub Initialise(lError As Long, szError As String, szRoutine As String)
  Set objError = New clsError
  With objError
    .Description = szError
    .Number = lError
    .Routine = szRoutine
  End With

  Me.Caption = App.Title & " Error"
  
  'show description error
  txtErr.Text = objError.Description & objError.Troubleshooting
    
  Me.Height = Me.Height - txtMore.Height - 60
  
  'center form
  Me.Top = (Screen.Height - Me.Height) / 2
  Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub cmdDetails_Click()
  If cmdDetails.Caption = "D&etails >>" Then
    Me.Height = Me.Height + txtMore.Height + 60
    cmdDetails.Caption = "<< D&etails"
    
    'show detail error
    cmdInfo_Click (4)
  Else
    Me.Height = Me.Height - txtMore.Height - 60
    cmdDetails.Caption = "D&etails >>"
  End If
End Sub
Private Sub cmdInfo_Click(Index As Integer)
Dim szTemp As String

  If Index = 0 Then
    txtMore.Text = objError.GetInfo(TIE_SYSTEM)
  ElseIf Index = 1 Then
    txtMore.Text = objError.GetInfo(TIE_APPLICATION)
  ElseIf Index = 2 Then
    txtMore.Text = objError.GetInfo(TIE_DATABASE)
  ElseIf Index = 3 Then
    txtMore.Text = objError.GetInfo(TIE_DRIVER_ODBC)
  ElseIf Index = 4 Then
    txtMore.Text = objError.GetInfo(TIE_ERROR)
  End If
End Sub

Private Sub cmdOK_Click()
  End
End Sub

Private Sub cmdSendMail_Click()
Dim szMail As String
Dim szTemp As String
Dim szSep As String

  szSep = String(60, "=")

  szTemp = objError.GetInfo(TIE_SYSTEM) & szSep & vbCrLf
  szTemp = szTemp & objError.GetInfo(TIE_APPLICATION) & szSep & vbCrLf
  szTemp = szTemp & objError.GetInfo(TIE_DATABASE) & szSep & vbCrLf
  szTemp = szTemp & objError.GetInfo(TIE_DRIVER_ODBC) & szSep & vbCrLf
  szTemp = szTemp & objError.GetInfo(TIE_ERROR) & String(60, "*") & vbCrLf
  szTemp = szTemp & "Insert your comment:" & vbCrLf

  szMail = "mailto:" & MailSupport & "?subject=Error Message&body=" & szTemp
  szMail = Replace(szMail, " ", "%20")
  szMail = Replace(szMail, Chr(10), "%0" & Hex(10))
  szMail = Replace(szMail, Chr(13), "%0" & Hex(13))
  szMail = Replace(szMail, Chr(32), "%0" & Hex(32))
  szMail = Replace(szMail, Chr(59), "%0" & Hex(59))
  
  'open shell
  ShellExecute hwnd, "open", szMail, vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub cmdContinue_Click()
  Unload Me
End Sub

