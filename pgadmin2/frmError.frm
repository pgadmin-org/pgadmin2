VERSION 5.00
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error"
   ClientHeight    =   5796
   ClientLeft      =   2880
   ClientTop       =   3156
   ClientWidth     =   7428
   Icon            =   "frmError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5796
   ScaleWidth      =   7428
   Begin VB.CheckBox chkAutoStart 
      Caption         =   "&Auto Start Application on exit"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkIgnoreError 
      Caption         =   "&Ignore this error (in this routine) until application is restarted"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   4455
   End
   Begin HighlightBox.TBX txtMore 
      Height          =   2535
      Left            =   1305
      TabIndex        =   10
      Top             =   3195
      Width           =   6045
      _ExtentX        =   10668
      _ExtentY        =   4466
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
   End
   Begin HighlightBox.TBX txtErr 
      Height          =   1296
      Left            =   48
      TabIndex        =   3
      Top             =   948
      Width           =   7308
      _ExtentX        =   12891
      _ExtentY        =   2265
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      ScaleHeight     =   2484
      ScaleWidth      =   1200
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
      Height          =   384
      Left            =   6720
      Picture         =   "frmError.frx":0442
      Top             =   240
      Width           =   384
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Error Handler"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
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
' This software is released under the Artistic Licence
'
' frmError.frm - Error form

' Note: No logging here

Option Explicit

Dim objError As clsError

Public Sub Initialise(lError As Long, szError As String, szRoutine As String, bSendMail As Boolean)
  Set objError = New clsError
  With objError
    .Description = szError
    .Number = lError
    .Routine = szRoutine
  End With

  'button send mail
  cmdSendMail.Enabled = bSendMail

  Me.Caption = App.Title & §§TrasLang§§(" Error")
  
  'show description error
  txtErr.Text = objError.Description & objError.Troubleshooting
    
  Me.Height = Me.Height - txtMore.Height - 60
  
  'center form
  Me.Top = (Screen.Height - Me.Height) / 2
  Me.Left = (Screen.Width - Me.Width) / 2
  
  PatchForm Me
End Sub

Private Sub cmdDetails_Click()
  If cmdDetails.Caption = §§TrasLang§§("D&etails >>") Then
    Me.Height = Me.Height + txtMore.Height + 60
    cmdDetails.Caption = §§TrasLang§§("<< D&etails")
    
    'show detail error
    cmdInfo_Click (4)
  Else
    Me.Height = Me.Height - txtMore.Height - 60
    cmdDetails.Caption = §§TrasLang§§("D&etails >>")
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
  'verify is execute a auto start application
  If chkAutoStart.Value = 1 Then
    Shell App.Path & "\" & App.EXEName & " " & Command, vbNormalFocus
  End If
  End
End Sub

Private Sub cmdSendMail_Click()
Dim szMail As String
Dim szTemp As String

  If frmMain.svr.LogLevel = llNone Then
    MsgBox §§TrasLang§§("Activate log error, Log settings are under Tools -> Options"), vbSystemModal + vbInformation, §§TrasLang§§("Activate Log")
    Exit Sub
  End If

  If Len(Trim(frmMain.svr.Logfile)) = 0 Then
    MsgBox §§TrasLang§§("Set log file name, Log settings are under Tools -> Options"), vbSystemModal + vbInformation, §§TrasLang§§("Activate Log")
    Exit Sub
  End If
  
  szTemp = objError.GetInfo(TIE_SYSTEM) & vbCrLf
  szTemp = szTemp & objError.GetInfo(TIE_APPLICATION) & vbCrLf
  szTemp = szTemp & objError.GetInfo(TIE_DATABASE) & vbCrLf
  szTemp = szTemp & objError.GetInfo(TIE_DRIVER_ODBC) & vbCrLf
  szTemp = szTemp & objError.GetInfo(TIE_ERROR) & vbCrLf & String(60, "*") & vbCrLf
  szTemp = szTemp & vbCrLf & "Insert your comment:" & vbCrLf

  szMail = "mailto:" & SUPPORT_EMAIL
  szMail = szMail & "?subject=pgA2 - Error Message: " & objError.Description
  szMail = szMail & "&body=" & szTemp
  
  szMail = Replace(szMail, " ", "%20")
  szMail = Replace(szMail, vbTab, "%0" & Hex(9))
  szMail = Replace(szMail, QUOTE, "%" & Hex(34))
  szMail = Replace(szMail, Chr(10), "%0" & Hex(10))
  szMail = Replace(szMail, Chr(13), "%0" & Hex(13))
  szMail = Replace(szMail, Chr(32), "%0" & Hex(32))
  szMail = Replace(szMail, Chr(59), "%0" & Hex(59))
  
  'open shell
  ShellExecute hwnd, "open", szMail, vbNullString, vbNullString, SW_SHOW
  
  MsgBox §§TrasLang§§("Add log file '") & frmMain.svr.Logfile & §§TrasLang§§("' to mail!"), vbInformation, §§TrasLang§§("Send Error Mail")
End Sub

Private Sub cmdContinue_Click()
  'store error in ingore error
  If chkIgnoreError.Value = vbChecked Then
    With objError
      ColIgnoreError.Add .Routine & "_" & .Number & "_" & .Description
    End With
  End If
  
  Unload Me
End Sub

