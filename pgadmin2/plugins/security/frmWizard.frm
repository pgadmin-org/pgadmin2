VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security Wizard"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList il 
      Left            =   540
      Top             =   3735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":08CA
            Key             =   "database"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":0A24
            Key             =   "group"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":0FBE
            Key             =   "public"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":1118
            Key             =   "sequence"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":16B2
            Key             =   "table"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":180C
            Key             =   "user"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":1DA6
            Key             =   "view"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picStrip 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      Picture         =   "frmWizard.frx":1F00
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   15
      Top             =   0
      Width           =   465
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   330
      Left            =   5445
      TabIndex        =   2
      ToolTipText     =   "Move back a stage"
      Top             =   3960
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   6480
      TabIndex        =   1
      ToolTipText     =   "Return SQL and exit."
      Top             =   3960
      Visible         =   0   'False
      Width           =   960
   End
   Begin TabDlg.SSTab tabWizard 
      Height          =   3840
      Left            =   495
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   45
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   6773
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   176
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmWizard.frx":2C2B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblInfo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvDatabases"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmWizard.frx":2C47
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblInfo(1)"
      Tab(1).Control(1)=   "lvObjects"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " "
      TabPicture(2)   =   "frmWizard.frx":2C63
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvEntities"
      Tab(2).Control(1)=   "lblInfo(2)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   " "
      TabPicture(3)   =   "frmWizard.frx":2C7F
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picContainer(0)"
      Tab(3).Control(1)=   "lblInfo(3)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   " "
      TabPicture(4)   =   "frmWizard.frx":2C9B
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picContainer(1)"
      Tab(4).Control(1)=   "lblInfo(6)"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   " "
      TabPicture(5)   =   "frmWizard.frx":2CB7
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "chkPermission(7)"
      Tab(5).Control(1)=   "chkPermission(6)"
      Tab(5).Control(2)=   "chkPermission(3)"
      Tab(5).Control(3)=   "chkPermission(5)"
      Tab(5).Control(4)=   "chkPermission(4)"
      Tab(5).Control(5)=   "chkPermission(2)"
      Tab(5).Control(6)=   "chkPermission(1)"
      Tab(5).Control(7)=   "chkPermission(0)"
      Tab(5).Control(8)=   "lblInfo(7)"
      Tab(5).ControlCount=   9
      TabCaption(6)   =   " "
      TabPicture(6)   =   "frmWizard.frx":2CD3
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblInfo(5)"
      Tab(6).Control(1)=   "lblInfo(4)"
      Tab(6).ControlCount=   2
      Begin VB.CheckBox chkPermission 
         Caption         =   "&Trigger"
         Height          =   195
         Index           =   7
         Left            =   -72435
         TabIndex        =   29
         Top             =   3285
         Width           =   1905
      End
      Begin VB.CheckBox chkPermission 
         Caption         =   "&References"
         Height          =   195
         Index           =   6
         Left            =   -72435
         TabIndex        =   28
         Top             =   2925
         Width           =   1905
      End
      Begin VB.CheckBox chkPermission 
         Caption         =   "&Delete"
         Height          =   195
         Index           =   3
         Left            =   -72435
         TabIndex        =   27
         Top             =   1845
         Width           =   1905
      End
      Begin VB.PictureBox picContainer 
         BorderStyle     =   0  'None
         Height          =   1635
         Index           =   1
         Left            =   -72435
         ScaleHeight     =   1635
         ScaleWidth      =   2490
         TabIndex        =   26
         Top             =   1440
         Width           =   2490
         Begin VB.OptionButton optClear 
            Caption         =   "&No, don't clear ACLs."
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   8
            Top             =   765
            Width           =   1905
         End
         Begin VB.OptionButton optClear 
            Caption         =   "&Yes, clear ACLs."
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   45
            Value           =   -1  'True
            Width           =   1635
         End
      End
      Begin VB.PictureBox picContainer 
         BorderStyle     =   0  'None
         Height          =   1410
         Index           =   0
         Left            =   -72615
         ScaleHeight     =   1410
         ScaleWidth      =   2310
         TabIndex        =   25
         Top             =   1305
         Width           =   2310
         Begin VB.OptionButton optAction 
            Caption         =   "&Grant Permissions"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   5
            Top             =   180
            Value           =   -1  'True
            Width           =   1635
         End
         Begin VB.OptionButton optAction 
            Caption         =   "&Revoke Permissions"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   6
            Top             =   900
            Width           =   1905
         End
      End
      Begin VB.CheckBox chkPermission 
         Caption         =   "&Rule"
         Height          =   195
         Index           =   5
         Left            =   -72435
         TabIndex        =   13
         Top             =   2565
         Width           =   1905
      End
      Begin VB.CheckBox chkPermission 
         Caption         =   "&Insert"
         Height          =   195
         Index           =   4
         Left            =   -72435
         TabIndex        =   12
         Top             =   2205
         Width           =   1905
      End
      Begin VB.CheckBox chkPermission 
         Caption         =   "&Update"
         Height          =   195
         Index           =   2
         Left            =   -72435
         TabIndex        =   11
         Top             =   1485
         Width           =   1905
      End
      Begin VB.CheckBox chkPermission 
         Caption         =   "&Select"
         Height          =   195
         Index           =   1
         Left            =   -72435
         TabIndex        =   10
         Top             =   1125
         Width           =   1905
      End
      Begin VB.CheckBox chkPermission 
         Caption         =   "&All"
         Height          =   195
         Index           =   0
         Left            =   -72435
         TabIndex        =   9
         Top             =   765
         Width           =   1905
      End
      Begin MSComctlLib.ListView lvDatabases 
         Height          =   2445
         Left            =   135
         TabIndex        =   0
         Top             =   1170
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   4313
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
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
            Text            =   "Database"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Comment"
            Object.Width           =   7939
         EndProperty
      End
      Begin MSComctlLib.ListView lvObjects 
         Height          =   2445
         Left            =   -74865
         TabIndex        =   3
         Top             =   1170
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   4313
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Object Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Database"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ACL"
            Object.Width           =   6174
         EndProperty
      End
      Begin MSComctlLib.ListView lvEntities 
         Height          =   2445
         Left            =   -74865
         TabIndex        =   4
         Top             =   1170
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   4313
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User or Group Name"
            Object.Width           =   11290
         EndProperty
      End
      Begin VB.Label lblInfo 
         Caption         =   "Select the permissions to Grant or Revoke."
         Height          =   825
         Index           =   7
         Left            =   -74820
         TabIndex        =   24
         Top             =   270
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   "Do you wish to clear down any existing Access Control Lists first?"
         Height          =   825
         Index           =   6
         Left            =   -74820
         TabIndex        =   23
         Top             =   270
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "All the information required has now been collected."
         Height          =   195
         Index           =   4
         Left            =   -73515
         TabIndex        =   22
         Top             =   1305
         Width           =   3645
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Click the OK button to apply your settings, or use the previous button to change them."
         Height          =   195
         Index           =   5
         Left            =   -74550
         TabIndex        =   21
         Top             =   2115
         Width           =   6045
      End
      Begin VB.Label lblInfo 
         Caption         =   "Select the users or groups that you wish to grant or revoke permissions to or from."
         Height          =   735
         Index           =   2
         Left            =   -74820
         TabIndex        =   20
         Top             =   270
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   "Select the objects for which you wish to update the Access Control List."
         Height          =   735
         Index           =   1
         Left            =   -74820
         TabIndex        =   19
         Top             =   270
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   "Do you wish to Grant or Revoke permissions?"
         Height          =   825
         Index           =   3
         Left            =   -74820
         TabIndex        =   18
         Top             =   270
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   "Select the databases containing the objects for which you wish to batch update the Access Control Lists (ACLs)."
         Height          =   735
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   270
         Width           =   6630
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   330
      Left            =   6480
      TabIndex        =   16
      ToolTipText     =   "Move forward a stage"
      Top             =   3960
      Width           =   960
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit
Dim bButtonPress As Boolean
Dim bProgramPress As Boolean

Private Sub chkPermission_Click(Index As Integer)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.chkPermission_Click()", etFullDebug

Dim X As Integer

  If Index = 0 Then
    If chkPermission(0).Value = 1 Then
      For X = 1 To 7
        chkPermission(X).Enabled = False
      Next X
    Else
      For X = 1 To 7
        chkPermission(X).Enabled = True
      Next X
    End If
  End If
  
  'Reset for < 7.2 if necessary
  If svr.dbVersion.VersionNum < 7.2 Then
    chkPermission(3).Enabled = False
    chkPermission(6).Enabled = False
    chkPermission(7).Enabled = False
  End If
  
  'Lock Update/Delete for PostgresQL < 7.2
  If ((Index = 2) And (svr.dbVersion.VersionNum < 7.2)) Then chkPermission(3).Value = chkPermission(2).Value
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.chkPermission_Click"
End Sub

Private Sub cmdNext_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdNext_Click()", etFullDebug

Dim objItem As ListItem

  bButtonPress = True
  Select Case tabWizard.Tab
    Case 0
      'Only move on if at least one database is selected.
      For Each objItem In lvDatabases.ListItems
        If objItem.Checked Then
          GetObjects
          tabWizard.Tab = 1
          cmdNext.Enabled = True
          cmdPrevious.Enabled = True
          Exit For
        End If
      Next objItem
    Case 1
      'Only move on if at least one object is selected.
      For Each objItem In lvObjects.ListItems
        If objItem.Checked Then
          GetEntities
          tabWizard.Tab = 2
          cmdNext.Enabled = True
          cmdPrevious.Enabled = True
          Exit For
        End If
      Next objItem
    Case 2
      'Only move on if at least one entity is selected.
      For Each objItem In lvEntities.ListItems
        If objItem.Checked Then
          tabWizard.Tab = 3
          cmdNext.Enabled = True
          cmdPrevious.Enabled = True
          Exit For
        End If
      Next objItem
    Case 3
      'If we are revoking permissions, then don't offer to revoke all first.
      If optAction(0).Value Then
        tabWizard.Tab = 4
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
      Else
        tabWizard.Tab = 5
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
      End If
    Case 4
      tabWizard.Tab = 5
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 5
      'Only move on if at least one permission has been selected
      If (chkPermission(0).Value = 1) Or (chkPermission(1).Value = 1) Or (chkPermission(2).Value = 1) Or (chkPermission(3).Value = 1) Or (chkPermission(4).Value = 1) Or (chkPermission(5).Value = 1) Or (chkPermission(6).Value = 1) Or (chkPermission(7).Value = 1) Then
        tabWizard.Tab = 6
        cmdNext.Enabled = False
        cmdNext.Visible = False
        cmdOK.Enabled = True
        cmdOK.Visible = True
        cmdPrevious.Enabled = True
      End If
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdNext_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.Form_Unload()", etFullDebug

  bRunning = False

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Form_Unload"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdOK_Click()", etFullDebug

Dim objOItem As ListItem
Dim objEItem As ListItem
Dim szEntity As String
Dim vEntity As Variant
Dim szUserList As String
Dim szAccessList As String
Dim szUsers() As String
Dim szAccess() As String
Dim lACL As Long

  StartMsg "Applying security..."
  
  For Each objOItem In lvObjects.ListItems
    If objOItem.Checked Then
      
      'Revoke all existing permissions if required.
      If (optAction(0).Value = True) And (optClear(0).Value = True) Then
        Select Case objOItem.Icon
          Case "sequence"
            ParseACL svr.Databases(objOItem.SubItems(1)).Sequences(objOItem.Text).ACL, szUserList, szAccessList
          Case "table"
            ParseACL svr.Databases(objOItem.SubItems(1)).Tables(objOItem.Text).ACL, szUserList, szAccessList
          Case "view"
            ParseACL svr.Databases(objOItem.SubItems(1)).Views(objOItem.Text).ACL, szUserList, szAccessList
        End Select
        szUsers = Split(szUserList, "|")
        szAccess = Split(szAccessList, "|")
        For Each vEntity In szUsers
          Select Case objOItem.Icon
            Case "sequence"
              If vEntity <> "" Then svr.Databases(objOItem.SubItems(1)).Sequences(objOItem.Text).Revoke vEntity, aclAll
            Case "table"
              If vEntity <> "" Then svr.Databases(objOItem.SubItems(1)).Tables(objOItem.Text).Revoke vEntity, aclAll
            Case "view"
              If vEntity <> "" Then svr.Databases(objOItem.SubItems(1)).Views(objOItem.Text).Revoke vEntity, aclAll
          End Select
        Next vEntity
      End If
    
      'Apply/Grant new permissions
      For Each objEItem In lvEntities.ListItems
        If objEItem.Checked Then
          If objEItem.Icon = "group" Then
            szEntity = "GROUP " & QUOTE & objEItem.Text & QUOTE
          ElseIf objEItem.Icon = "public" Then
            szEntity = "PUBLIC"
          Else
            szEntity = QUOTE & objEItem.Text & QUOTE
          End If
          lACL = 0
          If chkPermission(0).Value = 1 Then lACL = lACL + aclAll
          If chkPermission(1).Value = 1 Then lACL = lACL + aclSelect
          If chkPermission(2).Value = 1 Then lACL = lACL + aclUpdate
          If chkPermission(3).Value = 1 Then lACL = lACL + aclDelete
          If chkPermission(4).Value = 1 Then lACL = lACL + aclInsert
          If chkPermission(5).Value = 1 Then lACL = lACL + aclRule
          If chkPermission(6).Value = 1 Then lACL = lACL + aclReferences
          If chkPermission(7).Value = 1 Then lACL = lACL + aclTrigger
          If optAction(0).Value Then 'Grant permissions
            Select Case objOItem.Icon
              Case "sequence"
                svr.Databases(objOItem.SubItems(1)).Sequences(objOItem.Text).Grant szEntity, lACL
              Case "table"
                svr.Databases(objOItem.SubItems(1)).Tables(objOItem.Text).Grant szEntity, lACL
              Case "view"
                svr.Databases(objOItem.SubItems(1)).Views(objOItem.Text).Grant szEntity, lACL
            End Select
          Else 'Revoke permissions
            Select Case objOItem.Icon
              Case "sequence"
                svr.Databases(objOItem.SubItems(1)).Sequences(objOItem.Text).Revoke szEntity, lACL
              Case "table"
                svr.Databases(objOItem.SubItems(1)).Tables(objOItem.Text).Revoke szEntity, lACL
              Case "view"
                svr.Databases(objOItem.SubItems(1)).Views(objOItem.Text).Revoke szEntity, lACL
            End Select
          End If
        End If
      Next objEItem
    End If
  Next objOItem
  
  EndMsg
  bRunning = False
  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdOK_Click"
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdPrevious_Click()", etFullDebug

  bButtonPress = True
  Select Case tabWizard.Tab
    Case 6
      tabWizard.Tab = 5
      cmdNext.Enabled = True
      cmdNext.Visible = True
      cmdOK.Enabled = False
      cmdOK.Visible = False
      cmdPrevious.Enabled = True
    Case 5
      If optAction(0).Value Then
        tabWizard.Tab = 4
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
      Else
        tabWizard.Tab = 3
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
      End If
    Case 4
      tabWizard.Tab = 3
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 3
      tabWizard.Tab = 2
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 2
      tabWizard.Tab = 1
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 1
      tabWizard.Tab = 0
      cmdNext.Enabled = True
      cmdPrevious.Enabled = False
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdPrevious_Click"
End Sub

Private Sub tabWizard_Click(PreviousTab As Integer)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.tabWizard_Click(" & PreviousTab & ")", etFullDebug

  If bButtonPress = False And bProgramPress = False Then
    bProgramPress = True
    tabWizard.Tab = PreviousTab
  Else
    bProgramPress = False
  End If
  bButtonPress = False
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.tabWizard_Click"
End Sub

Public Sub Initialise()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.Initialise()", etFullDebug

Dim objDatabase As pgDatabase
Dim objItem As ListItem
  
  lvDatabases.ListItems.Clear
  tabWizard.Tab = 0
  cmdPrevious.Enabled = False
  
  StartMsg "Examining Server..."
  For Each objDatabase In svr.Databases
    If Not objDatabase.SystemObject Then
      Set objItem = lvDatabases.ListItems.Add(, , objDatabase.Identifier, "database", "database")
      objItem.SubItems(1) = Replace(objDatabase.Comment, vbCrLf, " ")
    End If
  Next objDatabase
  
  'Enable new permissions for versions of PostgreSQL >= 7.2
  If svr.dbVersion.VersionNum < 7.2 Then
    chkPermission(3).Enabled = False
    chkPermission(6).Enabled = False
    chkPermission(7).Enabled = False
  End If
  
  EndMsg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Initialise"
End Sub

Public Sub GetObjects()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.GetObjects()", etFullDebug

Dim objSequence As pgSequence
Dim objTable As pgTable
Dim objView As pgView
Dim objItem As ListItem
Dim objDBItem As ListItem
  
  lvObjects.ListItems.Clear
  
  StartMsg "Examining Server..."
  For Each objDBItem In lvDatabases.ListItems
    If objDBItem.Checked Then
    
      'Load Sequences
      For Each objSequence In svr.Databases(objDBItem.Text).Sequences
        If Not objSequence.SystemObject Then
          Set objItem = lvObjects.ListItems.Add(, , objSequence.Identifier, "sequence", "sequence")
          objItem.SubItems(1) = objSequence.Database
          objItem.SubItems(2) = objSequence.ACL
        End If
      Next objSequence
      
      'Load Tables
      For Each objTable In svr.Databases(objDBItem.Text).Tables
        If Not objTable.SystemObject Then
          Set objItem = lvObjects.ListItems.Add(, , objTable.Identifier, "table", "table")
          objItem.SubItems(1) = objTable.Database
          objItem.SubItems(2) = objTable.ACL
        End If
      Next objTable
          
      'Load Views
      For Each objView In svr.Databases(objDBItem.Text).Views
        If Not objView.SystemObject Then
          Set objItem = lvObjects.ListItems.Add(, , objView.Identifier, "view", "view")
          objItem.SubItems(1) = objView.Database
          objItem.SubItems(2) = objView.ACL
        End If
      Next objView
    End If
  Next objDBItem
  EndMsg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.GetObjects"
End Sub

Public Sub GetEntities()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.GetEntities()", etFullDebug

Dim objUser As pgUser
Dim objGroup As pgGroup
  
  lvEntities.ListItems.Clear
  lvEntities.ListItems.Add , , "PUBLIC", "public", "public"
  
  StartMsg "Examining Server..."
    For Each objUser In svr.Users
      lvEntities.ListItems.Add , , objUser.Identifier, "user", "user"
    Next objUser
    For Each objGroup In svr.Groups
      lvEntities.ListItems.Add , , objGroup.Identifier, "group", "group"
    Next objGroup
  EndMsg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.GetEntities"
End Sub
