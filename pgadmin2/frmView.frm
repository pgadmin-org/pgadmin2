VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   0
      Top             =   6390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3285
      TabIndex        =   4
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   5
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
      TabPicture(0)   =   "frmView.frx":06C2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboProperties(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "hbxProperties(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtProperties(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtProperties(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "&Definition"
      TabPicture(1)   =   "frmView.frx":06DE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "hbxProperties(1)"
      Tab(1).Control(1)=   "cmdLoad"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Security"
      TabPicture(2)   =   "frmView.frx":06FA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvProperties(0)"
      Tab(2).Control(1)=   "cmdRemove"
      Tab(2).Control(2)=   "cmdAdd"
      Tab(2).Control(3)=   "fraAdd"
      Tab(2).ControlCount=   4
      Begin VB.Frame fraAdd 
         Caption         =   "Define Privilege"
         Height          =   1815
         Left            =   -74865
         TabIndex        =   14
         Top             =   4410
         Width           =   5190
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Rule"
            Height          =   195
            Index           =   5
            Left            =   3420
            TabIndex        =   22
            ToolTipText     =   "Give rule privilege to the selected entity."
            Top             =   990
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Insert"
            Height          =   195
            Index           =   4
            Left            =   3420
            TabIndex        =   21
            ToolTipText     =   "Give insert privilege to the selected entity."
            Top             =   720
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Update"
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   19
            ToolTipText     =   "Give update privilege to the selected entity."
            Top             =   1260
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Select"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   18
            ToolTipText     =   "Give select privilege to the selected entity."
            Top             =   990
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&All"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   17
            ToolTipText     =   "Give all privileges to the selected entity."
            Top             =   720
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Delete"
            Height          =   195
            Index           =   3
            Left            =   225
            TabIndex        =   20
            ToolTipText     =   "Give delete privilege to the selected entity."
            Top             =   1530
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "R&eferences"
            Height          =   195
            Index           =   6
            Left            =   3420
            TabIndex        =   23
            ToolTipText     =   "Give references privilege to the selected entity."
            Top             =   1260
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Trigger"
            Height          =   195
            Index           =   7
            Left            =   3420
            TabIndex        =   24
            ToolTipText     =   "Give trigger privilege to the selected entity."
            Top             =   1530
            Width           =   1590
         End
         Begin MSComctlLib.ImageCombo cboEntities 
            Height          =   330
            Left            =   1260
            TabIndex        =   15
            ToolTipText     =   "Select a user, group or 'PUBLIC'."
            Top             =   315
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            ImageList       =   "il"
         End
         Begin VB.Label lblProperties 
            AutoSize        =   -1  'True
            Caption         =   "User/Group"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   16
            Top             =   360
            Width           =   840
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   -74865
         TabIndex        =   8
         ToolTipText     =   "Add the defined entry."
         Top             =   3915
         Width           =   1230
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   -73515
         TabIndex        =   9
         ToolTipText     =   "Remove the selected entry."
         Top             =   3915
         Width           =   1230
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   3390
         Index           =   0
         Left            =   -74865
         TabIndex        =   7
         ToolTipText     =   "The access control list for the view."
         Top             =   450
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   5980
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User/Group name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Privileges"
            Object.Width           =   4939
         EndProperty
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         Height          =   330
         Left            =   -74865
         TabIndex        =   13
         ToolTipText     =   "Load a query."
         Top             =   5895
         Width           =   945
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the view."
         Top             =   675
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The views OID (Object ID) in the PostgreSQL Database."
         Top             =   1080
         Width           =   3390
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   4245
         Index           =   0
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Comments about the view."
         Top             =   1935
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   7488
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
      Begin HighlightBox.HBX hbxProperties 
         Height          =   5325
         Index           =   1
         Left            =   -74865
         TabIndex        =   6
         ToolTipText     =   "The SQL query that will generate this view."
         Top             =   450
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   9393
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
         Caption         =   "View Definition (SQL)"
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   25
         ToolTipText     =   "The views owner."
         Top             =   1440
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
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
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   11
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   720
         Width           =   420
      End
   End
   Begin MSComctlLib.ImageList il 
      Left            =   540
      Top             =   6300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":0716
            Key             =   "user"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":0CB0
            Key             =   "group"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":124A
            Key             =   "public"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmView.frm - Edit/Create a View

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szNamespace As String
Dim szUsers() As String
Dim objView As pgView

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmView.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmView.cmdCancel_Click"
End Sub

Private Sub cmdLoad_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmView.cmdLoad_Click()", etFullDebug

Dim szLine As String
Dim szFile As String
Dim fNum As Integer
  
  With cdlg
    .DialogTitle = "Load SQL Query"
    .FLAGS = cdlOFNFileMustExist + cdlOFNHideReadOnly
    .Filter = "SQL Scripts (*.sql)|*.sql|All Files (*.*)|*.*"
    .FileName = ""
    .CancelError = True
    .ShowOpen
  End With
  
  If cdlg.FileName = "" Then Exit Sub
  hbxProperties(1).Text = ""
  fNum = FreeFile
  frmMain.svr.LogEvent "Loading " & cdlg.FileName, etMiniDebug
  Open cdlg.FileName For Input As #fNum
  While Not EOF(fNum)
    Line Input #fNum, szLine
    szFile = szFile & szLine & vbCrLf
  Wend
  If Len(szFile) > 2 Then szFile = Left(szFile, Len(szFile) - 2)
  
  Close #fNum
  hbxProperties(1).Text = szFile

  Exit Sub
Err_Handler:
  If Err.Number = 32755 Then
    frmMain.svr.LogEvent "Load Query operation cancelled.", etMiniDebug
    Exit Sub
  End If
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmView.cmdLoad_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmView.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim objNewView As pgView
Dim lACL As Long
Dim szEntity As String
Dim vEntity As Variant
Dim szComment As String
Dim szOldName As String
    
  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a view name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If UCase(Left(hbxProperties(1).Text, 6)) <> "SELECT" Then
    MsgBox "The view definition must start with 'SELECT'!", vbExclamation, "Error"
    tabProperties.Tab = 1
    hbxProperties(1).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating View..."
    Set objNewView = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views.Add(txtProperties(0).Text, hbxProperties(1).Text, hbxProperties(0).Text)
    
    'Add a new node and update the text on the parent
    Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views.Tag
    Set objNewView.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "VIE-" & GetID, txtProperties(0).Text, "view")
    objNode.Text = "Views (" & objNode.Children & ")"
    
  Else
    StartMsg "Updating View..."
    
    'Update the viewname if required
    If txtProperties(0).Tag = "Y" Then
      szOldName = objView.Name
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views.Rename szOldName, txtProperties(0).Text
        
      'Update the node text
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views(txtProperties(0).Text).Tag.Text = txtProperties(0).Text
    End If
    
    If hbxProperties(1).Tag = "Y" Then objView.Definition = hbxProperties(1).Text
    If hbxProperties(0).Tag = "Y" Then objView.Comment = hbxProperties(0).Text
  End If
  
  'Set the ACL on the View as required
  If lvProperties(0).Tag = "Y" Then
    'Revoke all from existing entries
    For Each vEntity In szUsers
      If vEntity <> "" Then
        If vEntity = "PUBLIC" Then
          frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views(txtProperties(0).Text).Revoke vEntity, aclAll
        ElseIf Left(vEntity, 6) = "GROUP " Then
          frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views(txtProperties(0).Text).Revoke "GROUP " & fmtID(Mid(vEntity, 7)), aclAll
        Else
          frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views(txtProperties(0).Text).Revoke fmtID(vEntity), aclAll
        End If
      End If
    Next vEntity
    
    'Now Grant the new permissions
    For Each objItem In lvProperties(0).ListItems
      If objItem.Icon = "group" Then
        szEntity = "GROUP " & fmtID(objItem.Text)
      ElseIf objItem.Icon = "public" Then
        szEntity = "PUBLIC"
      Else
        szEntity = fmtID(objItem.Text)
      End If
      lACL = 0
      If InStr(1, objItem.SubItems(1), "All") <> 0 Then lACL = lACL + aclAll
      If InStr(1, objItem.SubItems(1), "Select") <> 0 Then lACL = lACL + aclSelect
      If InStr(1, objItem.SubItems(1), "Update") <> 0 Then lACL = lACL + aclUpdate
      If InStr(1, objItem.SubItems(1), "Delete") <> 0 Then lACL = lACL + aclDelete
      If InStr(1, objItem.SubItems(1), "Insert") <> 0 Then lACL = lACL + aclInsert
      If InStr(1, objItem.SubItems(1), "Rule") <> 0 Then lACL = lACL + aclRule
      If InStr(1, objItem.SubItems(1), "References") <> 0 Then lACL = lACL + aclReferences
      If InStr(1, objItem.SubItems(1), "Trigger") <> 0 Then lACL = lACL + aclTrigger
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views(txtProperties(0).Text).Grant szEntity, lACL
    Next objItem
  End If
  
  'Finally, alter the username if required.
  If (cboProperties(0).Tag = "Y") And Not (frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views(txtProperties(0).Text).SystemObject) Then
    frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views(txtProperties(0).Text).Owner = cboProperties(0).Text
  End If
  
  'Simulate a node click to refresh the ListView
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmView.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, szNS As String, Optional View As pgView)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmView.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objItem As ListItem
Dim objUser As pgUser
Dim objGroup As pgGroup
Dim szUserlist As String
Dim szAccesslist As String
Dim szAccess() As String
  
  szDatabase = szDB
  szNamespace = szNS
  
  'Set the font
  For X = 0 To 1
    Set txtProperties(X).Font = ctx.Font
  Next X
  For X = 0 To 1
    Set hbxProperties(X).Font = ctx.Font
  Next X
  Set cboProperties(0).Font = ctx.Font
  Set cboEntities.Font = ctx.Font
  Set lvProperties(0).Font = ctx.Font
  hbxProperties(1).Wordlist = ctx.AutoHighlight
    
  For Each objUser In frmMain.svr.Users
    cboProperties(0).ComboItems.Add , objUser.Name, objUser.Name, "user"
  Next objUser
  
  'ACLs are different in 7.2+ and have 2 extra privileges
  If frmMain.svr.dbVersion.VersionNum < 7.2 Then
    chkPrivilege(6).Enabled = False
    chkPrivilege(7).Enabled = False
  End If
  
  If View Is Nothing Then
  
    'Create a new View
    bNew = True
    Me.Caption = "Create View"
    
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    cboProperties(0).BackColor = &H80000005
    hbxProperties(1).BackColor = &H80000005
    hbxProperties(1).Locked = False
    
    cboProperties(0).ComboItems(ctx.Username).Selected = True
        
    'Redim the userlist so it doesn't cause an error later.
    ReDim szUsers(0)
    
  Else
  
    'Display/Edit the specified View.
    Set objView = View
    bNew = False
    
    If objView.SystemObject Then  'Lock the permissions Add/Remove buttons if it's a system object
      cmdAdd.Enabled = False
      cmdRemove.Enabled = False
    End If
    
    If Not objView.SystemObject Then
      txtProperties(0).BackColor = &H80000005
      txtProperties(0).Locked = False
      cboProperties(0).BackColor = &H80000005
      hbxProperties(1).BackColor = &H80000005
      hbxProperties(1).Locked = False
    End If
    
    Me.Caption = "View: " & objView.Identifier
    txtProperties(0).Text = objView.Name
    txtProperties(1).Text = objView.OID
    If objView.SystemObject Then
      cboProperties(0).ComboItems.Clear
      cboProperties(0).ComboItems.Add , objView.Owner, objView.Owner, "user", "user"
    End If
    cboProperties(0).ComboItems(objView.Owner).Selected = True
    hbxProperties(0).Text = objView.Comment
    hbxProperties(1).Text = objView.Definition
    
    ParseACL objView.ACL, szUserlist, szAccesslist
    szUsers = Split(szUserlist, "|")
    szAccess = Split(szAccesslist, "|")
    For X = 0 To UBound(szUsers)
      If UCase(Left(szUsers(X), 6)) = "GROUP " Then
        Set objItem = lvProperties(0).ListItems.Add(, , Mid(szUsers(X), 7), "group", "group")
      Else
        If UCase(szUsers(X)) = "PUBLIC" Then
          Set objItem = lvProperties(0).ListItems.Add(, , szUsers(X), "public", "public")
        Else
          Set objItem = lvProperties(0).ListItems.Add(, , szUsers(X), "user", "user")
        End If
      End If
      objItem.SubItems(1) = szAccess(X)
    Next X
  End If
  
  'Load the Entities combo
  cboEntities.ComboItems.Add , , "PUBLIC", "public"
  For Each objUser In frmMain.svr.Users
    cboEntities.ComboItems.Add , , objUser.Name, "user"
  Next objUser
  For Each objGroup In frmMain.svr.Groups
    cboEntities.ComboItems.Add , , objGroup.Name, "group"
  Next objGroup
  cboEntities.ComboItems(1).Selected = True
  
  'Reset the Tags
  hbxProperties(0).Tag = "N"
  hbxProperties(1).Tag = "N"
  lvProperties(0).Tag = "N"
  txtProperties(0).Tag = "N"
  cboProperties(0).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmView.Initialise"
End Sub

Private Sub cmdRemove_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmView.cmdRemove_Click()", etFullDebug

  If lvProperties(0).SelectedItem Is Nothing Then Exit Sub
  lvProperties(0).ListItems.Remove lvProperties(0).SelectedItem.Index
  lvProperties(0).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmView.cmdRemove_Click"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmView.cmdAdd_Click()", etFullDebug

Dim szAccess As String
Dim objItem As ListItem

  If cboEntities.Text = "" Then Exit Sub
  
  'Check the entry doesn't already exist
  For Each objItem In lvProperties(0).ListItems
    If (objItem.Text = cboEntities.SelectedItem.Text) And (objItem.SmallIcon = cboEntities.SelectedItem.Image) Then
      MsgBox "'" & objItem.Text & "' already appears in the Access Control List. If you wish to modify this entry, it must be removed, and then replaced.", vbExclamation, "Error"
      Exit Sub
    End If
  Next objItem
  
  'Build the access string
  If chkPrivilege(0).Value = 1 Then
    szAccess = "All, "
  Else
    'ACLs are different in 7.2+
    If frmMain.svr.dbVersion.VersionNum < 7.2 Then
      If chkPrivilege(1).Value = 1 Then szAccess = szAccess & "Select, "
      If chkPrivilege(2).Value = 1 Then szAccess = szAccess & "Update/Delete, "
      If chkPrivilege(4).Value = 1 Then szAccess = szAccess & "Insert, "
      If chkPrivilege(5).Value = 1 Then szAccess = szAccess & "Rule, "
    Else
      If chkPrivilege(1).Value = 1 Then szAccess = szAccess & "Select, "
      If chkPrivilege(2).Value = 1 Then szAccess = szAccess & "Update, "
      If chkPrivilege(3).Value = 1 Then szAccess = szAccess & "Delete, "
      If chkPrivilege(4).Value = 1 Then szAccess = szAccess & "Insert, "
      If chkPrivilege(5).Value = 1 Then szAccess = szAccess & "Rule, "
      If chkPrivilege(6).Value = 1 Then szAccess = szAccess & "References, "
      If chkPrivilege(7).Value = 1 Then szAccess = szAccess & "Trigger, "
    End If
  End If
  If Len(szAccess) > 2 Then szAccess = Left(szAccess, Len(szAccess) - 2)
  If szAccess = "" Then szAccess = "None"
  
  Set objItem = lvProperties(0).ListItems.Add(, , cboEntities.SelectedItem.Text, cboEntities.SelectedItem.Image, cboEntities.SelectedItem.Image)
  objItem.SubItems(1) = szAccess
  lvProperties(0).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmView.cmdAdd_Click"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmView.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmView.hbxProperties_Change"
End Sub

Private Sub txtProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmView.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmView.txtProperties_Change"
End Sub

Private Sub chkPrivilege_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmView.chkPrivilege_Click(" & Index & ")", etFullDebug

Dim X As Integer

  If Index = 0 Then
    'ACLs are different in 7.2+
    If frmMain.svr.dbVersion.VersionNum < 7.2 Then
      If chkPrivilege(0).Value = 1 Then
        For X = 1 To 5
          chkPrivilege(X).Enabled = False
        Next X
      Else
        For X = 1 To 5
          chkPrivilege(X).Enabled = True
        Next X
      End If
    Else
      If chkPrivilege(0).Value = 1 Then
        For X = 1 To 7
          chkPrivilege(X).Enabled = False
        Next X
      Else
        For X = 1 To 7
          chkPrivilege(X).Enabled = True
        Next X
      End If
    End If
  End If
  
  'Link Update/Delete for older versions
  If frmMain.svr.dbVersion.VersionNum < 7.2 Then
    If Index = 2 Then chkPrivilege(3).Value = chkPrivilege(2).Value
    If Index = 3 Then chkPrivilege(2).Value = chkPrivilege(3).Value
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmView.chkPrivilege_Click"
End Sub

Private Sub cboProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmView.cboProperties_Click(" & Index & ")", etFullDebug

  cboProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmView.cboProperties_Click"
End Sub
