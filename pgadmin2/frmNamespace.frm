VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmNamespace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schema"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmNamespace.frx":0000
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
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      TabPicture(0)   =   "frmNamespace.frx":0BC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "hbxProperties(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtProperties(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtProperties(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboProperties(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "&Security"
      TabPicture(1)   =   "frmNamespace.frx":0BDE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraAdd"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAdd"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdRemove"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lvProperties(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   16
         Top             =   1170
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
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the Namespace."
         Top             =   375
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The Namespaces OID (Object ID) in the PostgreSQL Database."
         Top             =   780
         Width           =   3390
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   4245
         Index           =   0
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Comments about the Namespace."
         Top             =   1635
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
      Begin VB.Frame fraAdd 
         Caption         =   "Define Privilege"
         Height          =   1815
         Left            =   -74865
         TabIndex        =   9
         Top             =   4380
         Width           =   5190
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Create"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   11
            ToolTipText     =   "Give create privilege to the selected entity."
            Top             =   945
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Update"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   10
            ToolTipText     =   "Give update privilege to the selected entity."
            Top             =   1350
            Width           =   1590
         End
         Begin MSComctlLib.ImageCombo cboEntities 
            Height          =   330
            Left            =   1260
            TabIndex        =   12
            ToolTipText     =   "Select a user, group or 'PUBLIC'."
            Top             =   315
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Locked          =   -1  'True
            ImageList       =   "il"
         End
         Begin VB.Label lblProperties 
            AutoSize        =   -1  'True
            Caption         =   "User/Group"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   13
            Top             =   360
            Width           =   840
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74865
         TabIndex        =   14
         ToolTipText     =   "Add the defined entry."
         Top             =   3905
         Width           =   1230
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73515
         TabIndex        =   15
         ToolTipText     =   "Remove the selected entry."
         Top             =   3905
         Width           =   1230
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   3390
         Index           =   0
         Left            =   -74865
         TabIndex        =   17
         ToolTipText     =   "The access control list for the schema."
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
         BackColor       =   -2147483633
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
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   1230
         Width           =   465
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   825
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   420
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
            Picture         =   "frmNamespace.frx":0BFA
            Key             =   "user"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNamespace.frx":1194
            Key             =   "group"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNamespace.frx":172E
            Key             =   "public"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmNamespace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmNamespace.frm - Edit/Create a Namespace

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szUsers() As String
Dim objNamespace As pgNamespace

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmNamespace.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmNamespace.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmNamespace.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim objNewNamespace As pgNamespace
Dim lACL As Long
Dim szEntity As String
Dim vEntity As Variant
Dim szComment As String
    
  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Schema name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating Schema..."
    Set objNewNamespace = frmMain.svr.Databases(szDatabase).Namespaces.Add(txtProperties(0).Text, cboProperties(0).Text, hbxProperties(0).Text)
    
    'Add a new node and update the text on the parent
    Set objNode = frmMain.svr.Databases(szDatabase).Namespaces.Tag
    Set objNewNamespace.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "NSP-" & GetID, txtProperties(0).Text, "namespace")
    objNode.Text = "Schemas (" & objNode.Children & ")"
    
  Else
    StartMsg "Updating Schema..."
    If hbxProperties(0).Tag = "Y" Then objNamespace.Comment = hbxProperties(0).Text
  End If
  
  'Set the ACL on the Namespace as required
  If lvProperties(0).Tag = "Y" Then
    'Revoke all from existing entries
    For Each vEntity In szUsers
      If vEntity <> "" Then
        If vEntity = "PUBLIC" Then
          frmMain.svr.Databases(szDatabase).Namespaces(txtProperties(0).Text).Revoke vEntity, aclAll
        ElseIf Left(vEntity, 6) = "GROUP " Then
          frmMain.svr.Databases(szDatabase).Namespaces(txtProperties(0).Text).Revoke "GROUP " & fmtID(Mid(vEntity, 7)), aclAll
        Else
          frmMain.svr.Databases(szDatabase).Namespaces(txtProperties(0).Text).Revoke fmtID(vEntity), aclAll
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
      If InStr(1, objItem.SubItems(1), "Create") <> 0 Then lACL = lACL + aclCreate
      If InStr(1, objItem.SubItems(1), "Usage") <> 0 Then lACL = lACL + aclUsage
      frmMain.svr.Databases(szDatabase).Namespaces(txtProperties(0).Text).Grant szEntity, lACL
    Next objItem
  End If
  
  'Simulate a node click to refresh the ListNamespace
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmNamespace.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, Optional Namespace As pgNamespace)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmNamespace.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objItem As ListItem
Dim objCboItem As ComboItem
Dim objUser As pgUser
Dim objGroup As pgGroup
Dim szUserlist As String
Dim szAccesslist As String
Dim szAccess() As String
  
  szDatabase = szDB
  
  'Set the font
  For X = 0 To 1
    Set txtProperties(X).Font = ctx.Font
  Next X
  Set hbxProperties(0).Font = ctx.Font
  Set cboEntities.Font = ctx.Font
  Set lvProperties(0).Font = ctx.Font
  
  'Unlock the edittable fields
  If frmMain.svr.dbVersion.VersionNum >= 7.3 Then
    cmdAdd.Enabled = True
    cmdRemove.Enabled = True
    lvProperties(0).BackColor = &H80000005
    cboEntities.BackColor = &H80000005
    chkPrivilege(0).Enabled = True
    chkPrivilege(1).Enabled = True
  End If
  
  If Namespace Is Nothing Then
  
    'Create a new Namespace
    bNew = True
    Me.Caption = "Create Namespace"
    
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    cboProperties(0).BackColor = &H80000005
    cboProperties(0).Locked = False
    
    'Redim the userlist so it doesn't cause an error later.
    ReDim szUsers(0)
    
    'Populates the Owner's Combo & default to me
    cboProperties(0).ComboItems.Add , , ctx.Username, "user"
    For Each objUser In frmMain.svr.Users
      If objUser.Name <> ctx.Username Then cboProperties(0).ComboItems.Add , , objUser.Name, "user"
    Next objUser
    cboProperties(0).ComboItems(1).Selected = True
    
  Else
  
    'Display/Edit the specified Namespace.
    Set objNamespace = Namespace
    bNew = False
    
    If objNamespace.SystemObject Then  'Lock the permissions Add/Remove buttons if it's a system object
      cmdAdd.Enabled = False
      cmdRemove.Enabled = False
    End If
    
    Me.Caption = "Namespace: " & objNamespace.Identifier
    txtProperties(0).Text = objNamespace.Name
    txtProperties(1).Text = objNamespace.OID
    Set objCboItem = cboProperties(0).ComboItems.Add(, , objNamespace.Owner, "user")
    objCboItem.Selected = True
    hbxProperties(0).Text = objNamespace.Comment
    
    ParseACL objNamespace.ACL, szUserlist, szAccesslist
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
  If frmMain.svr.dbVersion.VersionNum >= 7.3 Then
    cboEntities.ComboItems.Add , , "PUBLIC", "public"
    For Each objUser In frmMain.svr.Users
      cboEntities.ComboItems.Add , , objUser.Name, "user"
    Next objUser
    For Each objGroup In frmMain.svr.Groups
      cboEntities.ComboItems.Add , , objGroup.Name, "group"
    Next objGroup
    cboEntities.ComboItems(1).Selected = True
  End If
  
  'Reset the Tags
  hbxProperties(0).Tag = "N"
  lvProperties(0).Tag = "N"
  txtProperties(0).Tag = "N"

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmNamespace.Initialise"
End Sub

Private Sub cmdRemove_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmNamespace.cmdRemove_Click()", etFullDebug

  If lvProperties(0).SelectedItem Is Nothing Then Exit Sub
  lvProperties(0).ListItems.Remove lvProperties(0).SelectedItem.Index
  lvProperties(0).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmNamespace.cmdRemove_Click"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmNamespace.cmdAdd_Click()", etFullDebug

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
  If chkPrivilege(0).Value = 1 Then szAccess = szAccess & "Create, "
  If chkPrivilege(1).Value = 1 Then szAccess = szAccess & "Usage, "
  If Len(szAccess) > 2 Then szAccess = Left(szAccess, Len(szAccess) - 2)
  If szAccess = "" Then szAccess = "None"
  
  Set objItem = lvProperties(0).ListItems.Add(, , cboEntities.SelectedItem.Text, cboEntities.SelectedItem.Image, cboEntities.SelectedItem.Image)
  objItem.SubItems(1) = szAccess
  lvProperties(0).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmNamespace.cmdAdd_Click"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmNamespace.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmNamespace.hbxProperties_Change"
End Sub

Private Sub txtProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmNamespace.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmNamespace.txtProperties_Change"
End Sub
