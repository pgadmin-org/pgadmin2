VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3285
      TabIndex        =   7
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   8
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
      TabPicture(0)   =   "frmDatabase.frx":014A
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
      Tab(0).Control(10)=   "txtProperties(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
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
         Index           =   3
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   $"frmDatabase.frx":0166
         Top             =   1890
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
         Height          =   3255
         Index           =   0
         Left            =   135
         TabIndex        =   6
         ToolTipText     =   "Comments about the database."
         Top             =   2745
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   5741
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
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   11
         Top             =   1530
         Width           =   465
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Encoding"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   1935
         Width           =   675
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Path"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   9
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
Dim objDatabase As pgDatabase

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmDatabase.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmDatabase.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmDatabase.cmdOK_Click()", etFullDebug

Dim objNode As Node

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a database name!", vbExclamation, "Error"
    txtProperties(0).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating Database..."
    frmMain.svr.Databases.Add txtProperties(0).Text, txtProperties(4).Text, txtProperties(3).Text, hbxProperties(0).Text
    
    'Add a new node and update the text on the parent
    For Each objNode In frmMain.tv.Nodes
      If Left(objNode.Key, 4) = "DAT+" Then
        frmMain.tv.Nodes.Add objNode.Key, tvwChild, "DAT-" & GetID, txtProperties(0).Text, "database"
        objNode.Text = "Databases (" & frmMain.svr.Databases.Count & ")"
      End If
    Next objNode
    
  Else
    StartMsg "Updating Database..."
    If hbxProperties(0).Tag = "Y" Then objDatabase.Comment = hbxProperties(0).Text
  End If
  
  'Simulate a node click to refresh the ListView
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmDatabase.cmdOK_Click"
End Sub

Public Sub Initialise(Optional Database As pgDatabase)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmDatabase.Initialise()", etFullDebug

Dim X As Integer
  
  If Database Is Nothing Then
  
    'Create a new database
    bNew = True
    Me.Caption = "Create Database"
    
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    txtProperties(3).BackColor = &H80000005
    txtProperties(3).Locked = False
    txtProperties(4).BackColor = &H80000005
    txtProperties(4).Locked = False
    
  Else
  
    'Display/Edit the specified Database.
    Set objDatabase = Database
    bNew = False
    Me.Caption = "Database: " & objDatabase.Identifier
    txtProperties(0).Text = objDatabase.Name
    txtProperties(1).Text = objDatabase.OID
    txtProperties(2).Text = objDatabase.Owner
    txtProperties(3).Text = objDatabase.Encoding
    txtProperties(4).Text = objDatabase.Path
    hbxProperties(0).Text = objDatabase.Comment
  End If
  
  'Reset the Tags
  For X = 0 To 4
    txtProperties(X).Tag = "N"
  Next X
  hbxProperties(0).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmDatabase.Initialise"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmDatabase.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmDatabase.hbxProperties_Change"
End Sub

Private Sub txtProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmDatabase.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmDatabase.txtProperties_Change"
End Sub
