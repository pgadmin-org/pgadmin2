VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRecordLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record query log"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmRecordLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Entry Types"
      Height          =   1410
      Left            =   270
      TabIndex        =   5
      Top             =   1035
      Width           =   4110
      Begin VB.CheckBox chkSystem 
         Caption         =   "System Queries"
         Height          =   195
         Left            =   1215
         TabIndex        =   8
         ToolTipText     =   "Record queries made by pgAdmin."
         Top             =   990
         Width           =   1500
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Datagrid Queries"
         Height          =   195
         Left            =   1215
         TabIndex        =   7
         ToolTipText     =   "Record queries generated whilst editting data in the datagrid."
         Top             =   675
         Width           =   1500
      End
      Begin VB.CheckBox chkUser 
         Caption         =   "User Queries"
         Height          =   195
         Left            =   1215
         TabIndex        =   6
         ToolTipText     =   "Record queries entered by the user."
         Top             =   360
         Width           =   1500
      End
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   90
      Top             =   2610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   330
      Left            =   3870
      TabIndex        =   3
      Top             =   495
      Width           =   330
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   495
      TabIndex        =   2
      ToolTipText     =   "Enter a filename to append log entries to."
      Top             =   495
      Width           =   3345
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2385
      TabIndex        =   1
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3555
      TabIndex        =   0
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Log file (entries will be appended to this file)"
      Height          =   195
      Left            =   495
      TabIndex        =   4
      Top             =   270
      Width           =   3060
   End
End
Attribute VB_Name = "frmRecordLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmRecordLog - Record an SQL Log.

Option Explicit

Private Sub cmdBrowse_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmRecordLog.cmdBrowse_Click()", etFullDebug
  
  With cdlg
    .FileName = txtFileName.Text
    .DialogTitle = "Record query log"
    .Filter = "SQL Scripts (*.sql)|*.sql|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    .CancelError = False
    .FLAGS = &H4
    .ShowOpen
  End With
  txtFileName.Text = cdlg.FileName

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmRecordLog.cmdBrowse_Click"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmRecordLog.cmdCancel_Click()", etFullDebug
  
  Unload Me

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmRecordLog.cmdCancel_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmRecordLog.Form_Load()", etFullDebug

  PatchForm Me
  
  txtFileName.Text = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "User Log Filename", frmMain.svr.UserLogfile)
  chkUser.Value = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Log User Queries", "0"))
  chkData.Value = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Log Data Queries", "0"))
  chkSystem.Value = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Log System Queries", "0"))
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmRecordLog.Form_Load"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmRecordLog.cmdOK_Click()", etFullDebug
  
Dim lLogOptions As Long

  If txtFileName.Text = "" Then
    MsgBox "You must enter or select a logfile!", vbExclamation, "Error"
    txtFileName.SetFocus
    Exit Sub
  End If
  
  'Get the log options
  If chkUser.Value = 1 Then lLogOptions = lLogOptions + qryUser
  If chkData.Value = 1 Then lLogOptions = lLogOptions + qryData
  If chkSystem.Value = 1 Then lLogOptions = lLogOptions + qrySystem
  
  If lLogOptions = 0 Then
    MsgBox "You must select at least on query type!", vbExclamation, "Error"
    chkUser.SetFocus
    Exit Sub
  End If
  
  frmMain.svr.UserLogfile = txtFileName.Text
  frmMain.svr.UserLogOptions = lLogOptions
  frmMain.svr.UserLog = True
  
  'Reset the menu
  frmMain.tb.Buttons("record").Enabled = False
  frmMain.tb.Buttons("stop").Enabled = True
  frmMain.mnuPopupRecordLog.Enabled = False
  frmMain.mnuPopupStopRecording = True
  
  'Save settings
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "User Log Filename", regString, txtFileName.Text
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Log User Queries", regString, chkUser.Value
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Log Data Queries", regString, chkData.Value
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Log System Queries", regString, chkSystem.Value
  
  Unload Me

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmRecordLog.cmdOK_Click"
End Sub
