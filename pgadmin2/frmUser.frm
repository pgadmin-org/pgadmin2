VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmUser.frx":0000
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
      TabPicture(0)   =   "frmUser.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtProperties(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtProperties(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtProperties(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtProperties(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fraPrivileges"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "mvProperties(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin MSComCtl2.MonthView mvProperties 
         Height          =   2370
         Index           =   0
         Left            =   2205
         TabIndex        =   7
         ToolTipText     =   "The date that the users account expires."
         Top             =   3780
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         ShowToday       =   0   'False
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   19791874
         CurrentDate     =   37089
         MinDate         =   36892
      End
      Begin VB.Frame fraPrivileges 
         Caption         =   "User Privileges"
         Height          =   1365
         Left            =   135
         TabIndex        =   14
         Top             =   2205
         Width           =   5190
         Begin VB.CheckBox chkProperties 
            Alignment       =   1  'Right Justify
            Caption         =   "Is this user a superuser?"
            Height          =   240
            Index           =   1
            Left            =   1035
            TabIndex        =   6
            ToolTipText     =   "Well, are they?"
            Top             =   900
            Width           =   2895
         End
         Begin VB.CheckBox chkProperties 
            Alignment       =   1  'Right Justify
            Caption         =   "Can this user create databases?"
            Height          =   240
            Index           =   0
            Left            =   1035
            TabIndex        =   5
            ToolTipText     =   "Well, can they?"
            Top             =   405
            Width           =   2895
         End
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the user."
         Top             =   540
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The user ID."
         Top             =   945
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1935
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "The users password."
         Top             =   1350
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1935
         PasswordChar    =   "*"
         TabIndex        =   4
         ToolTipText     =   "If you change the password, it will need to be re-entered here to confirm the changes."
         Top             =   1755
         Width           =   3390
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "User account expires"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   15
         Top             =   3780
         Width           =   1500
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Confirm password"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   13
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   12
         Top             =   1395
         Width           =   690
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "User ID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   11
         Top             =   990
         Width           =   540
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   585
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmUser.frm - Edit/Create a User

Option Explicit

Dim bNew As Boolean
Dim objUser As pgUser

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUser.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUser.cmdOK_Click()", etFullDebug

Dim objNode As Node

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a username!", vbExclamation, "Error"
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If txtProperties(2).Text <> txtProperties(3).Text Then
    MsgBox "The passwords do not match!", vbExclamation, "Error"
    txtProperties(2).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating User..."
    frmMain.svr.Users.Add txtProperties(0).Text, Val(txtProperties(1).Text), txtProperties(2).Text, Bin2Bool(chkProperties(0).Value), Bin2Bool(chkProperties(1).Value), mvProperties(0).Value
    
    'Add a new node and update the text on the parent
    For Each objNode In frmMain.tv.Nodes
      If Left(objNode.Key, 4) = "USR+" Then
        frmMain.tv.Nodes.Add objNode.Key, tvwChild, "USR-" & GetID, txtProperties(0).Text, "user"
        objNode.Text = "Users (" & frmMain.svr.Users.Count & ")"
      End If
    Next objNode
    
  Else
    StartMsg "Updating User..."
    If txtProperties(2).Tag = "Y" Then objUser.Password = txtProperties(2).Text
    If chkProperties(0).Tag = "Y" Then objUser.CreateDatabases = Bin2Bool(chkProperties(0).Value)
    If chkProperties(1).Tag = "Y" Then objUser.Superuser = Bin2Bool(chkProperties(1).Value)
    If mvProperties(0).Tag = "Y" Then objUser.AccountExpires = mvProperties(0).Value
  End If
  
  'Simulate a node click to refresh the ListView
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.cmdOK_Click"
End Sub

Public Sub Initialise(Optional User As pgUser)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUser.Initialise()", etFullDebug
  
Dim X As Integer
Dim objTempUser As pgUser
Dim lNextID As Long

  If User Is Nothing Then
  
    'Create a new user
    bNew = True
    Me.Caption = "Create User"
    
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    txtProperties(1).BackColor = &H80000005
    txtProperties(1).Locked = False
    
    'Set some defaults
    mvProperties(0).Value = DateAdd("yyyy", 1, Date)
    For Each objTempUser In frmMain.svr.Users
      If objTempUser.ID > lNextID Then lNextID = objTempUser.ID
    Next objTempUser
    txtProperties(1).Text = lNextID + 1
    
  Else
  
    'Display/Edit the specified User.
    Set objUser = User
    bNew = False
    Me.Caption = "User: " & objUser.Identifier
    txtProperties(0).Text = objUser.Name
    txtProperties(1).Text = objUser.ID
    txtProperties(2).Text = objUser.Password
    txtProperties(3).Text = objUser.Password
    chkProperties(0).Value = Bool2Bin(objUser.CreateDatabases)
    chkProperties(1).Value = Bool2Bin(objUser.Superuser)
    mvProperties(0).Value = objUser.AccountExpires
  End If
  
  'Reset the Tags
  For X = 0 To 3
    txtProperties(X).Tag = "N"
  Next X
  chkProperties(0).Tag = "N"
  chkProperties(1).Tag = "N"
  mvProperties(0).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.Initialise"
End Sub

Private Sub txtProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUser.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.txtProperties_Change"
End Sub

Private Sub chkProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUser.chkProperties_Click(" & Index & ")", etFullDebug

  If txtProperties(0).Text = "postgres" Then
    chkProperties(0).Value = Bool2Bin(objUser.CreateDatabases)
    chkProperties(1).Value = Bool2Bin(objUser.Superuser)
  Else
    chkProperties(Index).Tag = "Y"
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.chkProperties_Click"
End Sub

Private Sub mvProperties_DateClick(Index As Integer, ByVal DateClicked As Date)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUser.mvProperties_DateClick(" & Index & ")", etFullDebug

  mvProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.mvProperties_DateClick"
End Sub
