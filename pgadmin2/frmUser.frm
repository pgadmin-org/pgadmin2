VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
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
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      TabPicture(0)   =   "frmUser.frx":014A
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
      TabCaption(1)   =   "&Variables"
      TabPicture(1)   =   "frmUser.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRemoveVar"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAddVar"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtVarName"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtVarValue"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lvProperties(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdRemoveVar 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73515
         TabIndex        =   12
         ToolTipText     =   "Remove the selected variable."
         Top             =   4995
         Width           =   1230
      End
      Begin VB.CommandButton cmdAddVar 
         Caption         =   "&Add/Update"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74820
         TabIndex        =   11
         ToolTipText     =   "Add (or update) the specified variable."
         Top             =   4995
         Width           =   1230
      End
      Begin VB.TextBox txtVarName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73425
         TabIndex        =   13
         ToolTipText     =   "Enter the name of the variable to add or update."
         Top             =   5535
         Width           =   3750
      End
      Begin VB.TextBox txtVarValue 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73425
         TabIndex        =   14
         ToolTipText     =   "Enter the name of the variable to add or update."
         Top             =   5940
         Width           =   3750
      End
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
         StartOfWeek     =   61407234
         CurrentDate     =   37089
         MinDate         =   36892
      End
      Begin VB.Frame fraPrivileges 
         Caption         =   "User Privileges"
         Height          =   1365
         Left            =   135
         TabIndex        =   19
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
      Begin MSComctlLib.ListView lvProperties 
         Height          =   4425
         Index           =   0
         Left            =   -74865
         TabIndex        =   10
         ToolTipText     =   "Lists the configuration variables set for this user."
         Top             =   450
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   7805
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Variable"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Variable Name"
         Height          =   195
         Left            =   -74820
         TabIndex        =   22
         Top             =   5580
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "Variable Value"
         Height          =   195
         Left            =   -74820
         TabIndex        =   21
         Top             =   5985
         Width           =   1140
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "User account expires"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   20
         Top             =   3780
         Width           =   1500
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Confirm password"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   18
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   17
         Top             =   1395
         Width           =   690
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "User ID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   16
         Top             =   990
         Width           =   540
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   15
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
Dim szVarDropList As String

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
Dim objNewUser As pgUser
Dim objItem As ListItem
Dim szDropVars() As String
Dim X As Integer

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
    Set objNewUser = frmMain.svr.Users.Add(txtProperties(0).Text, Val(txtProperties(1).Text), txtProperties(2).Text, Bin2Bool(chkProperties(0).Value), Bin2Bool(chkProperties(1).Value), mvProperties(0).Value)
    
    'Add a new node and update the text on the parent
    Set objNode = frmMain.svr.Users.Tag
    Set objNewUser.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "USR-" & GetID, txtProperties(0).Text, "user")
    objNode.Text = "Users (" & frmMain.svr.Users.Count & ")"
      
  Else
    StartMsg "Updating User..."
    
    'Add any vars
    If lvProperties(0).Tag = "Y" Then
      For Each objItem In lvProperties(0).ListItems
        objUser.UserVars.AddOrUpdate objItem.Text, objItem.SubItems(1)
      Next objItem
    End If
    
    'Drop any vars
    If Len(szVarDropList) > 3 Then
      szDropVars = Split(szVarDropList, "!|!")
      For X = 0 To UBound(szDropVars)
        If szDropVars(X) <> "" Then objUser.UserVars.Remove szDropVars(X)
      Next X
    End If
    
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

  'If we error here, refresh the user vars to ensure they are in a consistant state
  If Not (objUser Is Nothing) Then
    objUser.UserVars.Refresh
    LoadVars
  End If
End Sub

Public Sub Initialise(Optional User As pgUser)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUser.Initialise()", etFullDebug
  
Dim X As Integer
Dim objTempUser As pgUser
Dim lNextID As Long

  'Set the font
  For X = 0 To 3
    Set txtProperties(X).Font = ctx.Font
  Next X
  Set txtVarValue.Font = ctx.Font
  Set txtVarName.Font = ctx.Font
  Set lvProperties(0).Font = ctx.Font
  
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
  
    'Unlock the Vars. We only edit these for existing objects as there is no
    'safe way to create the object & update the vars in one 'transaction'
    If ctx.dbVer >= 7.3 Then
      lvProperties(0).BackColor = &H80000005
      txtVarName.Enabled = True
      txtVarName.BackColor = &H80000005
      txtVarValue.Enabled = True
      txtVarValue.BackColor = &H80000005
      cmdAddVar.Enabled = True
      cmdRemoveVar.Enabled = True
    End If
  
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
    
    LoadVars
    
  End If
  
  'Reset the Tags
  For X = 0 To 3
    txtProperties(X).Tag = "N"
  Next X
  chkProperties(0).Tag = "N"
  chkProperties(1).Tag = "N"
  mvProperties(0).Tag = "N"
  lvProperties(0).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.Initialise"
End Sub

Private Sub LoadVars()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUser.LoadVars()", etFullDebug

Dim objItem As ListItem
Dim objVar As pgVar

  lvProperties(0).ListItems.Clear
  If ctx.dbVer >= 7.3 Then
    For Each objVar In objUser.UserVars
      Set objItem = lvProperties(0).ListItems.Add(, , objVar.Name)
      objItem.SubItems(1) = objVar.Value
    Next objVar
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.LoadVars"
End Sub

Private Sub cmdRemoveVar_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUser.cmdRemoveVar_Click()", etFullDebug

    If lvProperties(0).SelectedItem Is Nothing Then
    MsgBox "You must select a variable to remove!", vbExclamation, "Error"
    tabProperties.Tab = 1
    lvProperties(0).SetFocus
    Exit Sub
  End If
  
  If objUser Is Nothing Then
    lvProperties(0).ListItems.Remove lvProperties(0).SelectedItem.Index
    lvProperties(0).Tag = "Y"
    If lvProperties(0).SelectedItem Is Nothing Then
      cmdRemoveVar.Enabled = False
    End If
  Else
    szVarDropList = szVarDropList & lvProperties(0).SelectedItem.Text & "!|!"
    lvProperties(0).ListItems.Remove lvProperties(0).SelectedItem.Index
    lvProperties(0).Tag = "Y"
    If lvProperties(0).SelectedItem Is Nothing Then
      cmdRemoveVar.Enabled = False
    End If
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.cmdRemoveVar_Click"
End Sub

Private Sub cmdAddVar_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUser.cmdChkAdd_Click()", etFullDebug

Dim objItem As ListItem

  If txtVarName.Text = "" Then
    MsgBox "You must enter a name for the variable!", vbExclamation, "Error"
    tabProperties.Tab = 1
    txtVarName.SetFocus
    Exit Sub
  End If
  If txtVarValue.Text = "" Then
    MsgBox "You must enter a value for the variable!", vbExclamation, "Error"
    tabProperties.Tab = 1
    txtVarValue.SetFocus
    Exit Sub
  End If
  
  'Update
  For Each objItem In lvProperties(0).ListItems
    If objItem.Text = txtVarName.Text Then
      objItem.SubItems(1) = txtVarValue.Text
      lvProperties(0).Tag = "Y"
      txtVarName.Text = ""
      txtVarValue.Text = ""
      Exit Sub
    End If
  Next objItem
  
  'Or add
  Set objItem = lvProperties(0).ListItems.Add(, , txtVarName.Text)
  objItem.SubItems(1) = txtVarValue.Text
  lvProperties(0).Tag = "Y"
  
  txtVarName.Text = ""
  txtVarValue.Text = ""
  
  Exit Sub
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.cmdAddVar_Click"
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

Private Sub lvProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUser.lvProperties_Click(" & Index & ")", etFullDebug

  If Index = 0 Then
    If Not (lvProperties(0).SelectedItem Is Nothing) Then
      txtVarName.Text = lvProperties(0).SelectedItem.Text
      txtVarValue.Text = lvProperties(0).SelectedItem.SubItems(1)
    End If
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.lvProperties_Click"
End Sub

Private Sub mvProperties_DateClick(Index As Integer, ByVal DateClicked As Date)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUser.mvProperties_DateClick(" & Index & ")", etFullDebug

  mvProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.mvProperties_DateClick"
End Sub
