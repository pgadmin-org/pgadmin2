VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmGraveyard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Object Graveyard"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmGraveyard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtProperties 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   2
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Displays the parent table of the object (only applicable to Indexes, Rules and Triggers)."
      Top             =   900
      Width           =   3390
   End
   Begin VB.CommandButton cmdRestore 
      Cancel          =   -1  'True
      Caption         =   "&Restore"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2385
      TabIndex        =   11
      ToolTipText     =   "Restore the selected object."
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox txtProperties 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   5
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Displays the user who made the entry."
      Top             =   2115
      Width           =   3390
   End
   Begin VB.TextBox txtProperties 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   6
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Displays the action associated with this log entry."
      Top             =   2520
      Width           =   3390
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1215
      TabIndex        =   10
      ToolTipText     =   "Show the next log entry."
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      Height          =   375
      Left            =   45
      TabIndex        =   9
      ToolTipText     =   "Show the previous log entry."
      Top             =   6120
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   17
      Top             =   6570
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9622
            MinWidth        =   2822
         EndProperty
      EndProperty
   End
   Begin HighlightBox.HBX hbxProperties 
      Height          =   2040
      Index           =   0
      Left            =   45
      TabIndex        =   7
      ToolTipText     =   "Show the SQL definition of the object."
      Top             =   2925
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   3598
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
      Index           =   4
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Displays the version number of the object in the Revision Control log."
      Top             =   1710
      Width           =   3390
   End
   Begin VB.TextBox txtProperties 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   0
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Displays the object identifier."
      Top             =   90
      Width           =   3390
   End
   Begin VB.TextBox txtProperties 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Displays the object type."
      Top             =   495
      Width           =   3390
   End
   Begin VB.TextBox txtProperties 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   3
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Displays the timestamp of the object in the Revision Log."
      Top             =   1305
      Width           =   3390
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4365
      TabIndex        =   12
      ToolTipText     =   "Close this window."
      Top             =   6120
      Width           =   1095
   End
   Begin HighlightBox.HBX hbxProperties 
      Height          =   915
      Index           =   1
      Left            =   45
      TabIndex        =   8
      ToolTipText     =   "Displays comments about this version of the object."
      Top             =   5085
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   1614
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
      Caption         =   "Parent Table"
      Height          =   195
      Index           =   6
      Left            =   90
      TabIndex        =   20
      Top             =   945
      Width           =   915
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "User"
      Height          =   195
      Index           =   5
      Left            =   90
      TabIndex        =   19
      Top             =   2160
      Width           =   330
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Action"
      Height          =   195
      Index           =   4
      Left            =   90
      TabIndex        =   18
      Top             =   2565
      Width           =   450
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   16
      Top             =   1755
      Width           =   525
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Timestamp"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   15
      Top             =   1350
      Width           =   765
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Object Type"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   540
      Width           =   870
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Identifier"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   135
      Width           =   600
   End
End
Attribute VB_Name = "frmGraveyard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmGraveyard.frm - View object Graveyard from Revision Control Log

Option Explicit

Dim szDatabase As String
Dim lEntry As Long

Private Sub cmdClose_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmGraveyard.cmdClose_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmGraveyard.cmdClose_Click"
End Sub

Public Sub Initialise(szDB As String)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmGraveyard.Initialise(" & szDB & ")", etFullDebug

Dim szSQL As String

  szDatabase = szDB
  Me.Caption = "Object Graveyard: " & szDatabase
  hbxProperties(0).Wordlist = ctx.AutoHighlight
  
  'Disable buttons
  cmdNext.Enabled = True
  cmdPrevious.Enabled = True
  cmdRestore.Enabled = True
  
  If frmMain.svr.Databases(szDatabase).Graveyard.Count = 0 Then
    lEntry = 0
    sb.Panels(1).Text = "There are no objects in the graveyard."
    txtProperties(0).Text = ""
    txtProperties(1).Text = ""
    txtProperties(2).Text = ""
    txtProperties(3).Text = ""
    txtProperties(4).Text = ""
    txtProperties(5).Text = ""
    txtProperties(6).Text = ""
    hbxProperties(0).Text = ""
    hbxProperties(1).Text = ""
  Else
    lEntry = 1
    sb.Panels(1).Text = "Entry " & lEntry & " of " & frmMain.svr.Databases(szDatabase).Graveyard.Count
    txtProperties(0).Text = frmMain.svr.Databases(szDatabase).Graveyard(1).Identifier
    txtProperties(1).Text = frmMain.svr.Databases(szDatabase).Graveyard(1).ObjectType
    txtProperties(2).Text = frmMain.svr.Databases(szDatabase).Graveyard(1).ParentTable
    txtProperties(3).Text = frmMain.svr.Databases(szDatabase).Graveyard(1).TimeStamp
    txtProperties(4).Text = frmMain.svr.Databases(szDatabase).Graveyard(1).Version
    txtProperties(5).Text = frmMain.svr.Databases(szDatabase).Graveyard(1).User
    Select Case frmMain.svr.Databases(szDatabase).Graveyard(lEntry).Action
      Case "A"
        txtProperties(6).Text = "Object created (missing from database)."
      Case "U"
        txtProperties(6).Text = "Object updated (missing from database)."
      Case "D"
        txtProperties(6).Text = "Object deleted."
    End Select
    hbxProperties(0).Text = frmMain.svr.Databases(szDatabase).Graveyard(1).Definition
    hbxProperties(1).Text = frmMain.svr.Databases(szDatabase).Graveyard(1).Comment
    cmdRestore.Enabled = True
  End If
  
  'Set buttons
  If lEntry >= frmMain.svr.Databases(szDatabase).Graveyard.Count Then
    cmdNext.Enabled = False
  Else
    cmdNext.Enabled = True
  End If
  If lEntry <= 1 Then
    cmdPrevious.Enabled = False
  Else
    cmdPrevious.Enabled = True
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmGraveyard.Initialise"
End Sub

Private Sub cmdNext_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmGraveyard.cmdNext_Click()", etFullDebug

  lEntry = lEntry + 1
  sb.Panels(1).Text = "Entry " & lEntry & " of " & frmMain.svr.Databases(szDatabase).Graveyard.Count
  txtProperties(0).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).Identifier
  txtProperties(1).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).ObjectType
  txtProperties(2).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).ParentTable
  txtProperties(3).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).TimeStamp
  txtProperties(4).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).Version
  txtProperties(5).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).User
  Select Case frmMain.svr.Databases(szDatabase).Graveyard(lEntry).Action
    Case "A"
      txtProperties(6).Text = "Object created (missing from database)."
    Case "U"
      txtProperties(6).Text = "Object updated (missing from database)."
    Case "D"
      txtProperties(6).Text = "Object deleted."
  End Select
  hbxProperties(0).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).Definition
  hbxProperties(1).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).Comment
  
  If lEntry >= frmMain.svr.Databases(szDatabase).Graveyard.Count Then
    cmdNext.Enabled = False
  Else
    cmdNext.Enabled = True
  End If
  If lEntry <= 1 Then
    cmdPrevious.Enabled = False
  Else
    cmdPrevious.Enabled = True
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmGraveyard.cmdNext_Click"
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmGraveyard.cmdPrevious_Click()", etFullDebug

  lEntry = lEntry - 1
  sb.Panels(1).Text = "Entry " & lEntry & " of " & frmMain.svr.Databases(szDatabase).Graveyard.Count
  txtProperties(0).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).Identifier
  txtProperties(1).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).ObjectType
  txtProperties(2).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).ParentTable
  txtProperties(3).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).TimeStamp
  txtProperties(4).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).Version
  txtProperties(5).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).User
  Select Case frmMain.svr.Databases(szDatabase).Graveyard(lEntry).Action
    Case "A"
      txtProperties(6).Text = "Object created (missing from database)."
    Case "U"
      txtProperties(6).Text = "Object updated (missing from database)."
    Case "D"
      txtProperties(6).Text = "Object deleted."
  End Select
  hbxProperties(0).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).Definition
  hbxProperties(1).Text = frmMain.svr.Databases(szDatabase).Graveyard(lEntry).Comment
  
  If lEntry >= frmMain.svr.Databases(szDatabase).Graveyard.Count Then
    cmdNext.Enabled = False
  Else
    cmdNext.Enabled = True
  End If
  If lEntry <= 1 Then
    cmdPrevious.Enabled = False
  Else
    cmdPrevious.Enabled = True
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmGraveyard.cmdPrevious_Click"
End Sub

Private Sub cmdRestore_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmGraveyard.cmdRestore_Click()", etFullDebug

  If MsgBox("Are you sure you wish to restore the selected object?", vbQuestion + vbYesNo, "Restore Object?") = vbNo Then Exit Sub
  
  'Restore the object
  StartMsg "Restoring object..."
  frmMain.svr.Databases(szDatabase).Graveyard.Restore lEntry
   
  'Reset
  Initialise szDatabase
  EndMsg
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmGraveyard.cmdRestore_Click"
End Sub
