VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Object History"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRollback 
      Cancel          =   -1  'True
      Caption         =   "&Rollback"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2385
      TabIndex        =   14
      ToolTipText     =   "Show the next log entry."
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1215
      TabIndex        =   7
      ToolTipText     =   "Show the next log entry."
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      Height          =   375
      Left            =   45
      TabIndex        =   6
      ToolTipText     =   "Show the previous log entry."
      Top             =   6120
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Top             =   6570
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6764
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
      EndProperty
   End
   Begin HighlightBox.HBX hbxProperties 
      Height          =   3075
      Index           =   0
      Left            =   45
      TabIndex        =   4
      ToolTipText     =   "Show the SQL definition of the object."
      Top             =   1665
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   5424
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Object Definition"
   End
   Begin VB.TextBox txtProperties 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   3
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Displays the action associated with this log entry."
      Top             =   1260
      Width           =   3390
   End
   Begin VB.TextBox txtProperties 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   0
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Displays the entry timestamp."
      Top             =   45
      Width           =   3390
   End
   Begin VB.TextBox txtProperties 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Displays the version number for this entry."
      Top             =   450
      Width           =   3390
   End
   Begin VB.TextBox txtProperties 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   2
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Displays the user who made the entry."
      Top             =   855
      Width           =   3390
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4365
      TabIndex        =   8
      ToolTipText     =   "Close this window."
      Top             =   6120
      Width           =   1095
   End
   Begin HighlightBox.HBX hbxProperties 
      Height          =   1140
      Index           =   1
      Left            =   45
      TabIndex        =   5
      ToolTipText     =   "Displays comments about this version of the object."
      Top             =   4860
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   2011
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Comments"
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Action"
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   12
      Top             =   1305
      Width           =   450
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "User"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   11
      Top             =   900
      Width           =   330
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Top             =   495
      Width           =   525
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Timestamp"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   90
      Width           =   765
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmHistory.frm - View object history from Revision Control Log

Option Explicit

Dim objCurrent As Object
Dim lEntry As Long

Private Sub cmdClose_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmHistory.cmdClose_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmHistory.cmdClose_Click"
End Sub

Public Sub Initialise(objCurr As Object)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmHistory.Initialise(" & objCurr.Identifier & ")", etFullDebug

Dim szSQL As String

  Set objCurrent = objCurr
  Me.Caption = "Revision history: " & objCurrent.Identifier & " (" & objCurrent.ObjectType & ")"
  hbxProperties(0).Wordlist = ctx.AutoHighlight
  objCurrent.History.Refresh
  
  If objCurrent.History.Count = 0 Then
    lEntry = 0
    sb.Panels(1).Text = "This object is not in Revision Control."
  Else
    lEntry = 1
    sb.Panels(1).Text = "Entry " & objCurrent.History.Count & " of " & objCurrent.History.Count & " (shown most recent first)"
    txtProperties(0).Text = objCurrent.History(1).TimeStamp
    txtProperties(1).Text = objCurrent.History(1).Version
    txtProperties(2).Text = objCurrent.History(1).User
    Select Case objCurrent.History(lEntry).Action
      Case "A"
        txtProperties(3).Text = "Object created."
      Case "U"
        txtProperties(3).Text = "Object updated."
      Case "D"
        txtProperties(3).Text = "Object deleted."
    End Select
    hbxProperties(0).Text = objCurrent.History(1).Definition
    hbxProperties(1).Text = objCurrent.History(1).Comment
  End If
  
  'Enable the Previous button if required.
  If objCurrent.History.Count > 1 Then cmdPrevious.Enabled = True
  
  'Signal that this is the current version
  If hbxProperties(0).Text = objCurrent.SQL Then
    cmdRollback.Enabled = False
    sb.Panels(2).Bevel = sbrRaised
    sb.Panels(2).Text = "Current Version"
  Else
    cmdRollback.Enabled = True
    sb.Panels(2).Bevel = sbrInset
    sb.Panels(2).Text = "Previous Version"
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmHistory.Initialise"
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmHistory.cmdPrevious_Click()", etFullDebug

  lEntry = lEntry + 1
  sb.Panels(1).Text = "Entry " & objCurrent.History.Count - (lEntry - 1) & " of " & objCurrent.History.Count & " (shown most recent first)"
  txtProperties(0).Text = objCurrent.History(lEntry).TimeStamp
  txtProperties(1).Text = objCurrent.History(lEntry).Version
  txtProperties(2).Text = objCurrent.History(lEntry).User
  Select Case objCurrent.History(lEntry).Action
    Case "A"
      txtProperties(3).Text = "Object created."
    Case "U"
      txtProperties(3).Text = "Object updated."
    Case "D"
      txtProperties(3).Text = "Object deleted."
  End Select
  hbxProperties(0).Text = objCurrent.History(lEntry).Definition
  hbxProperties(1).Text = objCurrent.History(lEntry).Comment
  
  'Signal that this is the current version
  If hbxProperties(0).Text = objCurrent.SQL Then
    cmdRollback.Enabled = False
    sb.Panels(2).Bevel = sbrRaised
    sb.Panels(2).Text = "Current Version"
  Else
    cmdRollback.Enabled = True
    sb.Panels(2).Bevel = sbrInset
    sb.Panels(2).Text = "Previous Version"
  End If
  
  If lEntry >= objCurrent.History.Count Then
    cmdPrevious.Enabled = False
  Else
    cmdPrevious.Enabled = True
  End If
  If lEntry <= 1 Then
    cmdNext.Enabled = False
  Else
    cmdNext.Enabled = True
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmHistory.cmdPrevious_Click"
End Sub

Private Sub cmdNext_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmHistory.cmdNext_Click()", etFullDebug

  lEntry = lEntry - 1
  sb.Panels(1).Text = "Entry " & objCurrent.History.Count - (lEntry - 1) & " of " & objCurrent.History.Count & " (shown most recent first)"
  txtProperties(0).Text = objCurrent.History(lEntry).TimeStamp
  txtProperties(1).Text = objCurrent.History(lEntry).Version
  txtProperties(2).Text = objCurrent.History(lEntry).User
  Select Case objCurrent.History(lEntry).Action
    Case "A"
      txtProperties(3).Text = "Object created."
    Case "U"
      txtProperties(3).Text = "Object updated."
    Case "D"
      txtProperties(3).Text = "Object deleted."
  End Select
  hbxProperties(0).Text = objCurrent.History(lEntry).Definition
  hbxProperties(1).Text = objCurrent.History(lEntry).Comment
  
  'Signal that this is the current version
  If hbxProperties(0).Text = objCurrent.SQL Then
    cmdRollback.Enabled = False
    sb.Panels(2).Bevel = sbrRaised
    sb.Panels(2).Text = "Current Version"
  Else
    cmdRollback.Enabled = True
    sb.Panels(2).Bevel = sbrInset
    sb.Panels(2).Text = "Previous Version"
  End If
  
  If lEntry >= objCurrent.History.Count Then
    cmdPrevious.Enabled = False
  Else
    cmdPrevious.Enabled = True
  End If
  If lEntry <= 1 Then
    cmdNext.Enabled = False
  Else
    cmdNext.Enabled = True
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmHistory.cmdNext_Click"
End Sub

Private Sub cmdRollback_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmHistory.cmdNext_Click()", etFullDebug

  If objCurrent.ObjectType = "Table" Then
    If MsgBox("Are you sure you want to rollback this table to version " & txtProperties(1).Text & "? This will cause all data in this table to be lost, and any Indexes, Rules or Triggers will need to be restored from the Graveyard. Any Views or Functions that access this table may also be broken.", vbQuestion + vbYesNo, "Rollback Table") = vbNo Then Exit Sub
  Else
    If MsgBox("Are you sure you want to rollback this " & LCase(objCurrent.ObjectType) & " to version " & txtProperties(1).Text & "? Any other objects that refer to this object by it's OID may be broken.", vbQuestion + vbYesNo, "Rollback " & objCurrent.ObjectType) = vbNo Then Exit Sub
  End If
  
  StartMsg "Rolling back object..."
  Select Case objCurrent.ObjectType
    Case "Aggregate"
      frmMain.svr.Databases(objCurrent.Database).Aggregates.Rollback objCurrent.Identifier, txtProperties(1).Text
    Case "Function"
      frmMain.svr.Databases(objCurrent.Database).Functions.Rollback objCurrent.Identifier, txtProperties(1).Text
    Case "Index"
      frmMain.svr.Databases(objCurrent.Database).Tables(objCurrent.Table).Indexes.Rollback objCurrent.Identifier, txtProperties(1).Text
    Case "Language"
      frmMain.svr.Databases(objCurrent.Database).Languages.Rollback objCurrent.Identifier, txtProperties(1).Text
    Case "Operator"
      frmMain.svr.Databases(objCurrent.Database).Operators.Rollback objCurrent.Identifier, txtProperties(1).Text
    Case "Rule"
      frmMain.svr.Databases(objCurrent.Database).Tables(objCurrent.Table).Rules.Rollback objCurrent.Identifier, txtProperties(1).Text
    Case "Sequence"
      frmMain.svr.Databases(objCurrent.Database).Sequences.Rollback objCurrent.Identifier, txtProperties(1).Text
    Case "Table"
      frmMain.svr.Databases(objCurrent.Database).Tables.Rollback objCurrent.Identifier, txtProperties(1).Text
    Case "Trigger"
      frmMain.svr.Databases(objCurrent.Database).Tables(objCurrent.Table).Triggers.Rollback objCurrent.Identifier, txtProperties(1).Text
    Case "Type"
      frmMain.svr.Databases(objCurrent.Database).Types.Rollback objCurrent.Identifier, txtProperties(1).Text
    Case "View"
      frmMain.svr.Databases(objCurrent.Database).Views.Rollback objCurrent.Identifier, txtProperties(1).Text
  End Select
  
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
  Initialise objCurrent
  EndMsg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmHistory.cmdRollback_Click"
End Sub
