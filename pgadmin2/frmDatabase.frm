VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmDatabase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmDatabase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList il 
      Left            =   90
      Top             =   6210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":014A
            Key             =   "encoding"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3285
      TabIndex        =   8
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   9
      Top             =   6480
      Width           =   1095
   End
   Begin TabDlg.SSTab tabProperties 
      Height          =   6360
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   11218
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      TabPicture(0)   =   "frmDatabase.frx":0A24
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "hbxProperties(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtProperties(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtProperties(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtProperties(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProperties(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboProperties(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkProperties(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Revision Logging"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         ToolTipText     =   "Is Revision Logging enabled for this database? Once enabled, it can only be switched off by the database owner."
         Top             =   2745
         Width           =   2040
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   4
         ToolTipText     =   "Select or enter the encoding scheme to use."
         Top             =   1890
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "An alternate filesystem location in which to store the new database, specified as a string literal."
         Top             =   2295
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "The databases owner."
         Top             =   1485
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The databases OID (Object ID) in PostgreSQL."
         Top             =   1080
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the database."
         Top             =   675
         Width           =   3390
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   2895
         Index           =   0
         Left            =   135
         TabIndex        =   7
         ToolTipText     =   "Comments about the database."
         Top             =   3105
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   5106
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
         Locked          =   -1  'True
         Caption         =   "Comments"
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   14
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   13
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   12
         Top             =   1530
         Width           =   465
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Encoding"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   11
         Top             =   1935
         Width           =   675
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Path"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   10
         Top             =   2340
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmDatabase.frm - Edit/Create a Database

Option Explicit

Dim bNew As Boolean
Dim bSetting As Boolean
Dim objDatabase As pgDatabase

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.cmdOK_Click()", etFullDebug

Dim objNode As Node

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a database name!", vbExclamation, "Error"
    txtProperties(0).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating Database..."
    frmMain.svr.Databases.Add txtProperties(0).Text, txtProperties(3).Text, cboProperties(0).Text, hbxProperties(0).Text
    
    'Add a new node and update the text on the parent
    For Each objNode In frmMain.tv.Nodes
      If Left(objNode.Key, 4) = "DAT+" Then
        frmMain.tv.Nodes.Add objNode.Key, tvwChild, "DAT-" & GetID, txtProperties(0).Text, "database"
        objNode.Text = "Databases (" & objNode.Children & ")"
      End If
    Next objNode
    
  Else
    StartMsg "Updating Database..."
    If hbxProperties(0).Tag = "Y" Then objDatabase.Comment = hbxProperties(0).Text
  End If
  
  'Enable/Disable Revision Logging
  If chkProperties(0).Tag = "Y" Then
    frmMain.svr.Databases(txtProperties(0).Text).RevisionLogging = Bin2Bool(chkProperties(0).Value)
  End If
  
  'Simulate a node click to refresh the ListView
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.cmdOK_Click"
End Sub

Public Sub Initialise(Optional Database As pgDatabase)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.Initialise()", etFullDebug

Dim X As Integer
Dim objItem As ComboItem
  
  If Database Is Nothing Then
  
    'Create a new database
    bNew = True
    Me.Caption = "Create Database"
    
    'Load the Encoding Schemes
    cboProperties(0).ComboItems.Add , , "SQL_ASCII", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "EUC_JP", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "EUC_CN", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "EUC_KR", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "EUC_TW", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "UNICODE", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "MULE_INTERNAL", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "LATIN1", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "LATIN2", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "LATIN3", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "LATIN4", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "LATIN5", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "KOI8", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "WIN", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "ALT", "encoding", "encoding"
   
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    cboProperties(0).BackColor = &H80000005
    cboProperties(0).Locked = False
    txtProperties(3).BackColor = &H80000005
    txtProperties(3).Locked = False
    chkProperties(0).Enabled = False
    hbxProperties(0).BackColor = &H80000005
    hbxProperties(0).Locked = False
    
  Else
  
    'Display/Edit the specified Database.
    Set objDatabase = Database
    bNew = False
    Me.Caption = "Database: " & objDatabase.Identifier
    If objDatabase.Status <> statInaccessible Then
      chkProperties(0).Enabled = True
      hbxProperties(0).BackColor = &H80000005
      hbxProperties(0).Locked = False
    End If
    txtProperties(0).Text = objDatabase.Name
    txtProperties(1).Text = objDatabase.OID
    txtProperties(2).Text = objDatabase.Owner
    Set objItem = cboProperties(0).ComboItems.Add(, , objDatabase.EncodingName, "encoding", "encoding")
    objItem.Selected = True
    txtProperties(3).Text = objDatabase.Path
    bSetting = True
    chkProperties(0).Value = Bool2Bin(objDatabase.RevisionLogging)
    bSetting = False
    hbxProperties(0).Text = objDatabase.Comment
  End If
  
  'Reset the Tags
  For X = 0 To 3
    txtProperties(X).Tag = "N"
  Next X
  chkProperties(0).Tag = "N"
  hbxProperties(0).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.Initialise"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.hbxProperties_Change"
End Sub

Private Sub txtProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.txtProperties_Change"
End Sub

Private Sub chkProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.chkProperties_Click(" & Index & ")", etFullDebug

Dim bOrigSetting As Boolean

  If Not bSetting Then
    bOrigSetting = bSetting
    If Not (objDatabase Is Nothing) Then
      If (objDatabase.RevisionLogging) And (objDatabase.Owner <> ctx.Username) Then
        MsgBox "Only the database owner can switch off Revision Logging.", vbExclamation, "Error"
        bSetting = True
        chkProperties(0).Value = Bool2Bin(objDatabase.RevisionLogging)
        bSetting = bOrigSetting
        Exit Sub
      End If
      If (objDatabase.RevisionLogging) And (objDatabase.Owner = ctx.Username) Then
        If MsgBox("Switching of Revision Logging will delete the log table and all the Revision data it contains. Are you sure you wish to continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
          bSetting = True
          chkProperties(0).Value = Bool2Bin(objDatabase.RevisionLogging)
          bSetting = bOrigSetting
          Exit Sub
        End If
      End If
    End If
    chkProperties(0).Tag = "Y"
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.chkProperties_Click"
End Sub
