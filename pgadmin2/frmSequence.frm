VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmSequence 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sequence"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmSequence.frx":0000
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
      TabIndex        =   11
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   12
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
      TabPicture(0)   =   "frmSequence.frx":06C2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblProperties(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblProperties(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblProperties(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblProperties(9)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboProperties(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtProperties(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtProperties(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtProperties(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtProperties(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtProperties(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtProperties(6)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtProperties(7)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chkProperties(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "hbxProperties(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtProperties(5)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "&Security"
      TabPicture(1)   =   "frmSequence.frx":06DE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraAdd"
      Tab(1).Control(1)=   "cmdRemove"
      Tab(1).Control(2)=   "cmdAdd"
      Tab(1).Control(3)=   "lvProperties(0)"
      Tab(1).ControlCount=   4
      Begin VB.Frame fraAdd 
         Caption         =   "Define Privilege"
         Height          =   1815
         Left            =   -74865
         TabIndex        =   25
         Top             =   4410
         Width           =   5190
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Rule"
            Height          =   195
            Index           =   5
            Left            =   3420
            TabIndex        =   33
            ToolTipText     =   "Give rule privilege to the selected entity."
            Top             =   990
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Insert"
            Height          =   195
            Index           =   4
            Left            =   3420
            TabIndex        =   32
            ToolTipText     =   "Give insert privilege to the selected entity."
            Top             =   720
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Update"
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   30
            ToolTipText     =   "Give update privilege to the selected entity."
            Top             =   1260
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Select"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   29
            ToolTipText     =   "Give select privilege to the selected entity."
            Top             =   990
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&All"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   28
            ToolTipText     =   "Give all privileges to the selected entity."
            Top             =   720
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Delete"
            Height          =   195
            Index           =   3
            Left            =   225
            TabIndex        =   31
            ToolTipText     =   "Give delete privilege to the selected entity."
            Top             =   1530
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "R&eferences"
            Height          =   195
            Index           =   6
            Left            =   3420
            TabIndex        =   34
            ToolTipText     =   "Give references privilege to the selected entity."
            Top             =   1260
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Trigger"
            Height          =   195
            Index           =   7
            Left            =   3420
            TabIndex        =   35
            ToolTipText     =   "Give trigger privilege to the selected entity."
            Top             =   1530
            Width           =   1590
         End
         Begin MSComctlLib.ImageCombo cboEntities 
            Height          =   330
            Left            =   1260
            TabIndex        =   26
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
            Index           =   8
            Left            =   180
            TabIndex        =   27
            Top             =   360
            Width           =   840
         End
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   5
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "The initial starting value for the sequence."
         Top             =   3105
         Width           =   3390
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   -73515
         TabIndex        =   15
         ToolTipText     =   "Remove the selected entry."
         Top             =   3915
         Width           =   1230
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   -74865
         TabIndex        =   14
         ToolTipText     =   "Add the defined entry."
         Top             =   3915
         Width           =   1230
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   1455
         Index           =   0
         Left            =   135
         TabIndex        =   10
         ToolTipText     =   "Comments about the sequence."
         Top             =   4725
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   2566
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
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Cycled?"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   9
         ToolTipText     =   $"frmSequence.frx":06FA
         Top             =   4365
         Width           =   1995
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   7
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   $"frmSequence.frx":07F6
         Top             =   3915
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   6
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "A positive value will make an ascending sequence, a negative one a descending sequence."
         Top             =   3510
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "The maximum value for the sequence."
         Top             =   2700
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "The minimum value a sequence can generate."
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
         ToolTipText     =   "The last value of the sequence."
         Top             =   1890
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The sequences OID (Object ID) in the PostgreSQL Database."
         Top             =   1080
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         Height          =   285
         Index           =   0
         Left            =   1935
         TabIndex        =   1
         ToolTipText     =   "The name of the sequence."
         Top             =   675
         Width           =   3390
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   3390
         Index           =   0
         Left            =   -74865
         TabIndex        =   13
         ToolTipText     =   "The access control list for the sequence."
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
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   36
         ToolTipText     =   "The sequences owner."
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
         Caption         =   "Start value"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   24
         Top             =   3150
         Width           =   765
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Increment"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   23
         Top             =   3555
         Width           =   705
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Cache value"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   22
         Top             =   3960
         Width           =   900
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Last value"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   21
         Top             =   1935
         Width           =   735
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Minimum value"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   20
         Top             =   2340
         Width           =   1050
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Maximum value"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   19
         Top             =   2745
         Width           =   1095
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   17
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   16
         Top             =   1530
         Width           =   465
      End
   End
   Begin MSComctlLib.ImageList il 
      Left            =   0
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
            Picture         =   "frmSequence.frx":08A1
            Key             =   "user"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSequence.frx":0E3B
            Key             =   "group"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSequence.frx":13D5
            Key             =   "public"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmSequence.frm - Edit/Create a Sequence


Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szNamespace As String
Dim szUsers() As String
Dim objSequence As pgSequence

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSequence.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSequence.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSequence.cmdOK_Click()", etFullDebug

Dim szOldName As String
Dim objNode As Node
Dim objItem As ListItem
Dim objNewSequence As pgSequence
Dim lACL As Long
Dim szEntity As String
Dim vEntity As Variant

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Sequence name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If txtProperties(3).Tag = "Y" Then
    If MsgBox("Changing a Sequence value can be potentially dangerous, especially if that Sequence is used to create a Unique Key for a table. Are you sure you wish to continue?", vbQuestion + vbYesNo, "Change Value") = vbNo Then
      txtProperties(3).Text = objSequence.LastValue
      tabProperties.Tab = 0
      txtProperties(3).SetFocus
      Exit Sub
    End If
  End If
  
  'NOTE: Don't attempt to verify the sequence values - they're dependant on the architecture of the server OS
  '      so we'll just let PostgreSQL report any errors.
   
  If bNew Then
    StartMsg "Creating Sequence..."
    Set objNewSequence = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Sequences.Add(txtProperties(0).Text, txtProperties(6).Text, txtProperties(3).Text, txtProperties(4).Text, txtProperties(5).Text, txtProperties(7).Text, Bin2Bool(chkProperties(0).Value), hbxProperties(0).Text)
    
    'Add a new node and update the text on the parent
    Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Sequences.Tag
    Set objNewSequence.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "SEQ-" & GetID, txtProperties(0).Text, "sequence")
    objNode.Text = "Sequences (" & objNode.Children & ")"
    
  Else
    StartMsg "Updating Sequence..."
    
    'Update the sequencename if required
    If txtProperties(0).Tag = "Y" Then
      szOldName = objSequence.Name
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Sequences.Rename szOldName, txtProperties(0).Text
        
      'Update the node text
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Sequences(txtProperties(0).Text).Tag.Text = txtProperties(0).Text
    End If
    
    If txtProperties(2).Tag = "Y" Then objSequence.LastValue = txtProperties(2).Text
    If hbxProperties(0).Tag = "Y" Then objSequence.Comment = hbxProperties(0).Text
  End If
  
  'Set the ACL on the Sequence as required
  If lvProperties(0).Tag = "Y" Then
    'Revoke all from existing entries
    For Each vEntity In szUsers
      If vEntity <> "" Then
        If vEntity = "PUBLIC" Then
          frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Sequences(txtProperties(0).Text).Revoke vEntity, aclAll
        ElseIf Left(vEntity, 6) = "GROUP " Then
          frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Sequences(txtProperties(0).Text).Revoke "GROUP " & fmtID(Mid(vEntity, 7)), aclAll
        Else
          frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Sequences(txtProperties(0).Text).Revoke fmtID(vEntity), aclAll
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
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Sequences(txtProperties(0).Text).Grant szEntity, lACL
    Next objItem
  End If
  
  'Finally, alter the username if required.
  If (cboProperties(0).Tag = "Y") And Not (frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Sequences(txtProperties(0).Text).SystemObject) Then
    frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Sequences(txtProperties(0).Text).Owner = cboProperties(0).Text
  End If
  
  'Simulate a node click to refresh the ListSequence
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSequence.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, szNS As String, Optional Sequence As pgSequence)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSequence.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

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
  For X = 0 To 7
    Set txtProperties(X).Font = ctx.Font
  Next X
  Set cboProperties(0).Font = ctx.Font
  Set hbxProperties(0).Font = ctx.Font
  Set cboEntities.Font = ctx.Font
  Set lvProperties(0).Font = ctx.Font
  
  'ACLs are different in 7.2+ and have 2 extra privileges
  If ctx.dbVer < 7.2 Then
    chkPrivilege(6).Enabled = False
    chkPrivilege(7).Enabled = False
  End If
  
  For Each objUser In frmMain.svr.Users
    cboProperties(0).ComboItems.Add , "U~" & objUser.Name, objUser.Name, "user"
  Next objUser
  
  If Sequence Is Nothing Then
  
    'Create a new Sequence
    bNew = True
    Me.Caption = "Create Sequence"
    
    'Unlock the edittable fields
    cboProperties(0).BackColor = &H80000005
    For X = 3 To 7
      txtProperties(X).BackColor = &H80000005
      txtProperties(X).Locked = False
    Next X
    
    'Set some defaults
    txtProperties(3).Text = "1"
    If ctx.dbVer < 7.2 Then
      txtProperties(4).Text = "2147483647"
    Else
      txtProperties(4).Text = "9223372036854775807"
    End If
    txtProperties(5).Text = "1"
    txtProperties(6).Text = "1"
    txtProperties(7).Text = "1"
    cboProperties(0).ComboItems("U~" & ctx.Username).Selected = True
    
    'Redim the userlist so it doesn't cause an error later.
    ReDim szUsers(0)
    
  Else
  
    'Display/Edit the specified Sequence.
    Set objSequence = Sequence
    bNew = False
    Me.Caption = "Sequence: " & objSequence.Identifier
    
    If Not objSequence.SystemObject Then
      cboProperties(0).BackColor = &H80000005
      txtProperties(2).BackColor = &H80000005
      txtProperties(2).Locked = False
    Else 'Lock the permissions Add/Remove buttons if it's a system object
      cmdAdd.Enabled = False
      cmdRemove.Enabled = False
    End If
    
    txtProperties(0).Text = objSequence.Name
    txtProperties(1).Text = objSequence.OID
    txtProperties(2).Text = objSequence.LastValue
    txtProperties(3).Text = objSequence.Minimum
    txtProperties(4).Text = objSequence.Maximum
    txtProperties(6).Text = objSequence.Increment
    txtProperties(7).Text = objSequence.Cache
    If objSequence.SystemObject Then
      cboProperties(0).ComboItems.Clear
      cboProperties(0).ComboItems.Add , objSequence.Owner, objSequence.Owner, "user", "user"
    End If
    cboProperties(0).ComboItems(objSequence.Owner).Selected = True
    chkProperties(0).Value = Bool2Bin(objSequence.Cycled)
    hbxProperties(0).Text = objSequence.Comment
    
    ParseACL objSequence.ACL, szUserlist, szAccesslist
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
  For X = 0 To 7
    txtProperties(X).Tag = "N"
  Next X
  cboProperties(0).Tag = "N"
  hbxProperties(0).Tag = "N"
  lvProperties(0).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSequence.Initialise"
End Sub

Private Sub cmdRemove_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSequence.cmdRemove_Click()", etFullDebug

  If lvProperties(0).SelectedItem Is Nothing Then Exit Sub
  lvProperties(0).ListItems.Remove lvProperties(0).SelectedItem.Index
  lvProperties(0).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSequence.cmdRemove_Click"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSequence.cmdAdd_Click()", etFullDebug

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
    If ctx.dbVer < 7.2 Then
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
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSequence.cmdAdd_Click"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSequence.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSequence.hbxProperties_Change"
End Sub

Private Sub txtProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSequence.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSequence.txtProperties_Change"
End Sub

Private Sub chkPrivilege_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSequence.chkPrivilege_Click(" & Index & ")", etFullDebug

Dim X As Integer

  If Index = 0 Then
    'ACLs are different in 7.2+
    If ctx.dbVer < 7.2 Then
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
  If ctx.dbVer < 7.2 Then
    If Index = 2 Then chkPrivilege(3).Value = chkPrivilege(2).Value
    If Index = 3 Then chkPrivilege(2).Value = chkPrivilege(3).Value
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSequence.chkPrivilege_Click"
End Sub

Private Sub chkProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSequence.chkProperties_Click(" & Index & ")", etFullDebug

  If Not (objSequence Is Nothing) Then
    chkProperties(0).Value = Bool2Bin(objSequence.Cycled)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUser.chkProperties_Click"
End Sub

Private Sub cboProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmSequence.cboProperties_Click(" & Index & ")", etFullDebug

  cboProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSequence.cboProperties_Click"
End Sub
