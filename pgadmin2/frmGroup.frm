VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmGroup.frx":0000
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
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      TabPicture(0)   =   "frmGroup.frx":06C2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtProperties(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtProperties(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lvProperties(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "il"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin MSComctlLib.ImageList il 
         Left            =   90
         Top             =   5670
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
               Picture         =   "frmGroup.frx":06DE
               Key             =   "user"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   4830
         Index           =   0
         Left            =   1890
         TabIndex        =   3
         ToolTipText     =   "All PostgreSQL users are listed, those that are ticked are members of the group."
         Top             =   1350
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   8520
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
            Text            =   "Username"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "User ID"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The user group ID."
         Top             =   945
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the user group."
         Top             =   540
         Width           =   3390
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Members"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   1395
         Width           =   645
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Group name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   585
         Width           =   870
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Group ID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   990
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmGroup.frm - Edit/Create a Group

Option Explicit

Dim bNew As Boolean
Dim objGroup As pgGroup

Private Sub cmdCancel_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmGroup.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmGroup.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmGroup.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim objNewGroup As pgGroup
Dim szAddList As String
Dim szRemoveList As String
  
  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a group name!", vbExclamation, "Error"
    txtProperties(0).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating Group..."
    Set objNewGroup = frmMain.svr.Groups.Add(txtProperties(0).Text, Val(txtProperties(1).Text))

    'Add a new node and update the text on the parent
    On Error Resume Next
    Set objNode = frmMain.svr.Groups.Tag
    Set objNewGroup.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "GRP-" & GetID, txtProperties(0).Text, "group")
    objNode.Text = "Groups (" & frmMain.svr.Groups.Count & ")"
    If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
    
  Else
    StartMsg "Updating Group..."
  End If
  
  'Add/Remove the users from the existing/new group
  For Each objItem In lvProperties(0).ListItems
    If objItem.Tag <> objItem.Checked Then 'Item has changed
      If objItem.Checked Then
        frmMain.svr.Groups(txtProperties(0).Text).Members.Add objItem.Text
      Else
        frmMain.svr.Groups(txtProperties(0).Text).Members.Remove objItem.Text
      End If
    End If
  Next objItem
  
  'Simulate a node click to refresh the ListView
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmGroup.cmdOK_Click"
End Sub

Public Sub Initialise(Optional Group As pgGroup)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmGroup.Initialise()", etFullDebug
  
Dim X As Integer
Dim objItem As ListItem
Dim objTempGroup As pgGroup
Dim objTempMember As Variant
Dim objTempUser As pgUser
Dim lNextID As Long

  PatchForm Me
  
  'Load the users...
  For Each objTempUser In frmMain.svr.Users
    Set objItem = lvProperties(0).ListItems.Add(, "U:" & objTempUser.Name, objTempUser.Name, "user", "user")
    objItem.SubItems(1) = objTempUser.ID
    objItem.Tag = False
  Next objTempUser
    
  If Group Is Nothing Then
  
    'Create a new Group
    bNew = True
    Me.Caption = "Create Group"
    
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    txtProperties(1).BackColor = &H80000005
    txtProperties(1).Locked = False
    
    'Set defaults
    For Each objTempGroup In frmMain.svr.Groups
      If objTempGroup.ID > lNextID Then lNextID = objTempGroup.ID
    Next objTempGroup
    txtProperties(1).Text = lNextID + 1
    
  Else
  
    'Display/Edit the specified Group.
    Set objGroup = Group
    bNew = False
    Me.Caption = "Group: " & objGroup.Identifier
    txtProperties(0).Text = objGroup.Name
    txtProperties(1).Text = objGroup.ID

    'Tick the included users. Note that instead of using the item tag to mark that the item has
    'been changed, we use it to store the original value. This will solve errors caused when items
    'are mistakenly checked, then unchecked.
    For Each objTempMember In Group.Members
      lvProperties(0).ListItems("U:" & objTempMember).Checked = True
      lvProperties(0).ListItems("U:" & objTempMember).Tag = True
    Next objTempMember
  End If
  
  'Reset the Tags
  txtProperties(0).Tag = "N"
  txtProperties(1).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmGroup.Initialise"
End Sub

Private Sub lvProperties_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmGroup.lvProperties_ItemCheck(" & Index & ", " & Item.Text & ")", etFullDebug

  Item.Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmGroup.lvProperties_ItemCheck"
End Sub

Private Sub txtProperties_Change(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmGroup.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmGroup.txtProperties_Change"
End Sub


