VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User"
   ClientHeight    =   6885
   ClientLeft      =   7770
   ClientTop       =   1875
   ClientWidth     =   5520
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5520
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
      Tab(0).Control(5)=   "il"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtProperties(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtProperties(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtProperties(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProperties(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fraPrivileges"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "mvProperties(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "&Variables"
      TabPicture(1)   =   "frmUser.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(2)=   "cboVarName"
      Tab(1).Control(3)=   "lvProperties(0)"
      Tab(1).Control(4)=   "txtVarValue"
      Tab(1).Control(5)=   "cmdAddVar"
      Tab(1).Control(6)=   "cmdRemoveVar"
      Tab(1).Control(7)=   "cboVarValue"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "&Present in groups"
      TabPicture(2)   =   "frmUser.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvProperties(1)"
      Tab(2).ControlCount=   1
      Begin MSComctlLib.ImageCombo cboVarValue 
         Height          =   330
         Left            =   -73425
         TabIndex        =   24
         Top             =   5940
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
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
      Begin VB.TextBox txtVarValue 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73425
         TabIndex        =   13
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
         StartOfWeek     =   58851330
         CurrentDate     =   37089
         MinDate         =   36892
      End
      Begin VB.Frame fraPrivileges 
         Caption         =   "User Privileges"
         Height          =   1365
         Left            =   135
         TabIndex        =   18
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
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
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
      Begin MSComctlLib.ListView lvProperties 
         Height          =   5745
         Index           =   1
         Left            =   -74865
         TabIndex        =   22
         ToolTipText     =   "Lists the configuration variables set for this user."
         Top             =   450
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   10134
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSComctlLib.ImageList il 
         Left            =   360
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":019E
               Key             =   "group"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":0870
               Key             =   "property"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":0F42
               Key             =   "on"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":1394
               Key             =   "off"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":17E6
               Key             =   "warning"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":1C38
               Key             =   "error"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":208A
               Key             =   "info"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":24DC
               Key             =   "debug"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUser.frx":2A76
               Key             =   "log"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageCombo cboVarName 
         Height          =   330
         Left            =   -73425
         TabIndex        =   23
         Top             =   5520
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin VB.Label Label1 
         Caption         =   "Variable Name"
         Height          =   195
         Left            =   -74820
         TabIndex        =   21
         Top             =   5580
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "Variable Value"
         Height          =   195
         Left            =   -74820
         TabIndex        =   20
         Top             =   5985
         Width           =   1140
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "User account expires"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   19
         Top             =   3780
         Width           =   1500
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Confirm password"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   17
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   16
         Top             =   1395
         Width           =   690
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "User ID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   15
         Top             =   990
         Width           =   540
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   14
         Top             =   585
         Width           =   720
      End
   End
   Begin VB.Menu mnuModify 
      Caption         =   "Modify"
      Visible         =   0   'False
      Begin VB.Menu mnuModifyCopyVar 
         Caption         =   "Copy Setting Variable"
      End
      Begin VB.Menu mnuModifyPasteVar 
         Caption         =   "Paste Setting Variable"
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmUser.frm - Edit/Create a User

Option Explicit

Dim bNew As Boolean
Dim objUser As pgUser
Dim szVarDropList As String
Const PrefKey = "KEY_"

Private Sub cboVarName_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUser.cboVarName_Click()", etFullDebug

Dim objVardb As VarDb
Dim vData
Dim szImg As String

  txtVarValue.Visible = False
  txtVarValue.Text = ""
  cboVarValue.Visible = False
  cboVarValue.Text = ""
  cboVarValue.ComboItems.Clear
  cboVarValue.Locked = True
  
  objVardb = GetVarDb(cboVarName.Text)
  
  If objVardb.Type = TVDB_BOOLEAN Then
    szImg = GetImageFromVal("ON", TVDB_BOOLEAN)
    cboVarValue.ComboItems.Add , PrefKey & "on", "ON", szImg, szImg
    szImg = GetImageFromVal("OFF", TVDB_BOOLEAN)
    cboVarValue.ComboItems.Add , PrefKey & "off", "OFF", szImg, szImg
    cboVarValue.ComboItems(1).Selected = True
    cboVarValue.Locked = True
    cboVarValue.Visible = True
  ElseIf objVardb.Type = TVDB_FLOAT Or objVardb.Type = TVDB_INTEGR Or objVardb.Type = TVDB_STRING Then
    txtVarValue.Visible = True
  ElseIf objVardb.Type = TVDB_CAST Then
    For Each vData In objVardb.CastValue
      szImg = GetImageFromVal(CStr(vData), TVDB_CAST)
      cboVarValue.ComboItems.Add , PrefKey & LCase(vData), vData, szImg, szImg
    Next
    cboVarValue.ComboItems(1).Selected = True
    cboVarValue.Locked = False
    cboVarValue.Visible = True
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.cmdCancel_Click"
End Sub

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
Dim objTempGroup As pgGroup
Dim objTempMember As Variant
Dim lNextID As Long
Dim objItem As ListItem
Dim rsVar As Recordset

  PatchForm Me
  
  lvProperties(0).ListItems.Clear

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
      txtVarValue.Enabled = True
      txtVarValue.BackColor = &H80000005
      cmdAddVar.Enabled = True
      cmdRemoveVar.Enabled = True
      cboVarName.Enabled = True
      cboVarName.BackColor = &H80000005
      cboVarValue.Enabled = True
      cboVarValue.BackColor = &H80000005
    
      'load var name
      cboVarName.ComboItems.Clear
      Set rsVar = frmMain.svr.Databases(frmMain.svr.MasterDB).Execute("SELECT name FROM pg_settings ORDER BY name")
      While Not rsVar.EOF
        cboVarName.ComboItems.Add , LCase(rsVar("name")), rsVar("name"), "property"
        rsVar.MoveNext
      Wend
      cboVarName.ComboItems(1).Selected = True
      cboVarName_Click
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
    
    'Present in groups
    For Each objTempGroup In frmMain.svr.Groups
      For Each objTempMember In objTempGroup.Members
        If objTempMember = objUser.Name Then
          Set objItem = lvProperties(1).ListItems.Add(, , objTempGroup.Name, "group", "group")
        End If
      Next objTempMember
    Next objTempGroup
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
Dim objVardb As VarDb
Dim szImg As String

  lvProperties(0).ListItems.Clear
  If ctx.dbVer >= 7.3 Then
    For Each objVar In objUser.UserVars
      Set objItem = lvProperties(0).ListItems.Add(, , objVar.Name)
      objItem.SubItems(1) = objVar.Value
      
      'get image
      szImg = "property"    'image default
      objVardb = GetVarDb(objVar.Name)
      If objVardb.Type = TVDB_BOOLEAN Or objVardb.Type = TVDB_CAST Then
        szImg = GetImageFromVal(objVar.Value, objVardb.Type)
      End If
      objItem.Icon = szImg
      objItem.SmallIcon = szImg
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
Dim szVal As String
Dim szImg As String

  'get value variable
  If txtVarValue.Visible = True Then
    szVal = txtVarValue.Text
  ElseIf cboVarValue.Visible = True Then
    szVal = cboVarValue.Text
  End If

  If Trim(szVal) = "" Then
    MsgBox "You must enter a value for the variable!", vbExclamation, "Error"
    tabProperties.Tab = 1
    txtVarValue.SetFocus
    Exit Sub
  End If
 
  'image default
  szImg = "property"
  
  'Update
  For Each objItem In lvProperties(0).ListItems
    If objItem.Text = cboVarName.SelectedItem.Text Then
      objItem.SubItems(1) = szVal
      lvProperties(0).Tag = "Y"
      
      If cboVarValue.Visible And Not cboVarValue.SelectedItem Is Nothing Then
        If Len(cboVarValue.SelectedItem.Image) > 0 Then
          szImg = cboVarValue.SelectedItem.Image
        End If
      End If
      objItem.Icon = szImg
      objItem.SmallIcon = szImg
      
      cboVarName.ComboItems(1).Selected = True
      cboVarName_Click
      Exit Sub
    End If
  Next objItem
  
  'Or add
  Set objItem = lvProperties(0).ListItems.Add(, , cboVarName.SelectedItem.Text)
  objItem.SubItems(1) = szVal
  lvProperties(0).Tag = "Y"
  
  If cboVarValue.Visible And Not cboVarValue.SelectedItem Is Nothing Then
    If Len(cboVarValue.SelectedItem.Image) > 0 Then
      szImg = cboVarValue.SelectedItem.Image
    End If
  End If
  objItem.Icon = szImg
  objItem.SmallIcon = szImg
  
  cboVarName.ComboItems(1).Selected = True
  cboVarName_Click
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

Dim objVardb As VarDb
Dim vData
Dim szVal As String

  If Index = 0 Then
    If Not (lvProperties(0).SelectedItem Is Nothing) Then
      cboVarName.ComboItems(LCase(lvProperties(0).SelectedItem.Text)).Selected = True
      cboVarName_Click
      
      If txtVarValue.Visible = True Then
        txtVarValue.Text = lvProperties(0).SelectedItem.SubItems(1)
      ElseIf cboVarValue.Visible = True Then
        objVardb = GetVarDb(cboVarName.Text)
        If objVardb.Type = TVDB_BOOLEAN Then
          Select Case UCase(lvProperties(0).SelectedItem.SubItems(1))
            Case "ON", "TRUE", "YES", "1"
              szVal = "ON"
            Case "OFF", "FALSE", "NO", "0"
              szVal = "OFF"
            Case Else
              szVal = "OFF"
          End Select
          cboVarValue.ComboItems(PrefKey & LCase(szVal)).Selected = True
        Else
          szVal = lvProperties(0).SelectedItem.SubItems(1)
        End If
        
        'find value in combo
        For Each vData In objVardb.CastValue
          If UCase(szVal) = UCase(vData) Then
            cboVarValue.ComboItems(PrefKey & LCase(szVal)).Selected = True
            Exit Sub
          End If
        Next
        
        'cast value not present
        'manual insert
        If objVardb.Type = TVDB_CAST Then cboVarValue.Text = szVal
      End If
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


Private Sub lvProperties_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.lvProperties_MouseDown(" & Index & "," & Button & "," & Shift & "," & X & "," & y & ")", etFullDebug

  If Button = vbRightButton Then
    mnuModifyPasteVar.Enabled = False
    If ColVarDbBuffer.Count > 0 Then mnuModifyPasteVar.Enabled = True
    PopupMenu mnuModify
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.lvProperties_MouseDown"
End Sub

'copy var setting database
Private Sub mnuModifyCopyVar_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.mnuModifyCopyVar_Click()", etFullDebug

Dim objLv As ListItem

  Set ColVarDbBuffer = New Collection
  For Each objLv In lvProperties(0).ListItems
    ColVarDbBuffer.Add objLv.Text & "|" & objLv.SubItems(1)
  Next
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.mnuModifyCopyVar_Click"
End Sub

'paste var setting database
Private Sub mnuModifyPasteVar_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.mnuModifyPasteVar_Click()", etFullDebug

Dim vData

  'simulate add/update var
  For Each vData In ColVarDbBuffer
    vData = Split(vData, "|")
    'select variable name
    cboVarName.ComboItems(vData(0)).Selected = True
    cboVarName_Click
    
    'set value
    If cboVarValue.Visible Then
      cboVarValue.ComboItems(PrefKey & LCase(vData(1))).Selected = True
    Else
      txtVarValue.Text = vData(1)
    End If
    
    'add var
    cmdAddVar_Click
  Next
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.mnuModifyPasteVar_Click"
End Sub


