VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSysConf Wizard"
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
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":08CA
            Key             =   "database"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picStrip 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      Picture         =   "frmWizard.frx":0A24
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   8
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   45
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   6773
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   176
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmWizard.frx":18A6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblInfo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvDatabases"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdDatabaseNone"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdDatabaseAll"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmWizard.frx":18C2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optPasswordCaching(1)"
      Tab(1).Control(1)=   "optPasswordCaching(0)"
      Tab(1).Control(2)=   "lblInfo(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   " "
      TabPicture(2)   =   "frmWizard.frx":18DE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "udPopulationDelay"
      Tab(2).Control(1)=   "txtPopulationDelay"
      Tab(2).Control(2)=   "lblInfo(2)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   " "
      TabPicture(3)   =   "frmWizard.frx":18FA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtPopulationSize"
      Tab(3).Control(1)=   "udPopulationSize"
      Tab(3).Control(2)=   "lblInfo(3)"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   " "
      TabPicture(4)   =   "frmWizard.frx":1916
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblInfo(5)"
      Tab(4).Control(1)=   "lblInfo(4)"
      Tab(4).ControlCount=   2
      Begin VB.CommandButton cmdDatabaseAll 
         Height          =   555
         Left            =   6345
         Picture         =   "frmWizard.frx":1932
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Select all databases"
         Top             =   1755
         Width           =   555
      End
      Begin VB.CommandButton cmdDatabaseNone 
         Height          =   555
         Left            =   6345
         Picture         =   "frmWizard.frx":21FC
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Deselect all databases"
         Top             =   2430
         Width           =   555
      End
      Begin VB.TextBox txtPopulationSize 
         Height          =   285
         Left            =   -72345
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "100"
         Top             =   2115
         Width           =   1020
      End
      Begin MSComCtl2.UpDown udPopulationDelay 
         Height          =   285
         Left            =   -71310
         TabIndex        =   5
         Top             =   2115
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   393216
         Value           =   10
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtPopulationDelay"
         BuddyDispid     =   196613
         OrigLeft        =   3645
         OrigTop         =   2205
         OrigRight       =   3840
         OrigBottom      =   2490
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtPopulationDelay 
         Height          =   285
         Left            =   -72345
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "10"
         Top             =   2115
         Width           =   1005
      End
      Begin VB.OptionButton optPasswordCaching 
         Caption         =   "&Yes, allow the local caching of passwords."
         Height          =   240
         Index           =   1
         Left            =   -73515
         TabIndex        =   4
         Top             =   2475
         Width           =   3570
      End
      Begin VB.OptionButton optPasswordCaching 
         Caption         =   "&No, do not allow the local caching of passwords."
         Height          =   240
         Index           =   0
         Left            =   -73515
         TabIndex        =   3
         Top             =   1710
         Value           =   -1  'True
         Width           =   3750
      End
      Begin MSComctlLib.ListView lvDatabases 
         Height          =   2445
         Left            =   135
         TabIndex        =   0
         Top             =   1170
         Width           =   6135
         _ExtentX        =   10821
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Database"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "MSysConf Present?"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Local Passwords?"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Delay"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Size"
            Object.Width           =   1323
         EndProperty
      End
      Begin MSComCtl2.UpDown udPopulationSize 
         Height          =   285
         Left            =   -71310
         TabIndex        =   6
         Top             =   2115
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   393216
         Value           =   100
         BuddyControl    =   "txtPopulationSize"
         BuddyDispid     =   196612
         OrigLeft        =   3645
         OrigTop         =   2205
         OrigRight       =   3840
         OrigBottom      =   2490
         Increment       =   10
         Max             =   10000
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Click the OK button to apply your settings, or use the previous button to change them."
         Height          =   195
         Index           =   5
         Left            =   -74550
         TabIndex        =   17
         Top             =   2025
         Width           =   6045
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "All the information required has now been collected."
         Height          =   195
         Index           =   4
         Left            =   -73515
         TabIndex        =   16
         Top             =   1440
         Width           =   3645
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmWizard.frx":2AC6
         Height          =   825
         Index           =   3
         Left            =   -74820
         TabIndex        =   14
         Top             =   270
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmWizard.frx":2B93
         Height          =   825
         Index           =   2
         Left            =   -74820
         TabIndex        =   12
         Top             =   270
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmWizard.frx":2CBC
         Height          =   735
         Index           =   1
         Left            =   -74820
         TabIndex        =   11
         Top             =   270
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmWizard.frx":2D76
         Height          =   735
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   270
         Width           =   6630
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   330
      Left            =   6480
      TabIndex        =   9
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

Private Sub cmdDatabaseAll_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdDatabaseAll_Click()", etFullDebug

Dim objItem As ListItem

  For Each objItem In lvDatabases.ListItems
    objItem.Checked = True
  Next objItem
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdDatabaseAll_Click"
End Sub

Private Sub cmdDatabaseNone_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdDatabaseNone_Click()", etFullDebug

Dim objItem As ListItem

  For Each objItem In lvDatabases.ListItems
    objItem.Checked = False
  Next objItem
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdDatabaseNone_Click"
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
          tabWizard.Tab = 1
          cmdNext.Enabled = True
          cmdPrevious.Enabled = True
          Exit For
        End If
      Next objItem
    Case 1
      tabWizard.Tab = 2
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 2
      tabWizard.Tab = 3
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 3
      tabWizard.Tab = 4
      cmdNext.Enabled = False
      cmdNext.Visible = False
      cmdOK.Enabled = True
      cmdOK.Visible = True
      cmdPrevious.Enabled = True
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

Dim objItem As ListItem

  StartMsg "Configuring MSysConf tables..."
  For Each objItem In lvDatabases.ListItems
    If objItem.Checked Then
    
      'Create table if required.
      If objItem.SubItems(1) = "No" Then
        svr.LogEvent "Creating MSysConf table in " & objItem.Text, etMiniDebug
        svr.Databases(objItem.Text).Tables.Add "msysconf", "config int4 NOT NULL, chvalue varchar(255), nvalue int4, comments varchar(255)", , , , , "The MSysConf table contains global settings for the Microsoft Jet Engine."
        svr.Databases(objItem.Text).Tables("msysconf").Grant "PUBLIC", aclSelect
      End If
      
      'Drop all existing records before reinserting. This is easier than figuring out
      'if they exist already or not to determine whether to insert or update them.
      svr.Databases(objItem.Text).Execute "DELETE FROM msysconf"
      
      'Now insert all three records.
      If Not optPasswordCaching(0).Value Then
        svr.Databases(objItem.Text).Execute "INSERT INTO msysconf VALUES ('101', '', '1', 'Allow local storage of passwords.')"
      Else
        svr.Databases(objItem.Text).Execute "INSERT INTO msysconf VALUES ('101', '', '0', 'Disallow local storage of passwords.')"
      End If
      svr.Databases(objItem.Text).Execute "INSERT INTO msysconf VALUES ('102', '', '" & Val(txtPopulationDelay.Text) & "', 'Background population delay = " & Val(txtPopulationDelay.Text) & " seconds.')"
      svr.Databases(objItem.Text).Execute "INSERT INTO msysconf VALUES ('103', '', '" & Val(txtPopulationSize.Text) & "', 'Background population size = " & Val(txtPopulationSize.Text) & " records.')"
      
    End If
  Next objItem
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
    Case 4
      tabWizard.Tab = 3
      cmdNext.Enabled = True
      cmdNext.Visible = True
      cmdOK.Enabled = False
      cmdOK.Visible = False
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
svr.LogEvent "Entering " & App.Title & ":frmSQLWizard.tabWizard_Click(" & PreviousTab & ")", etFullDebug

  If bButtonPress = False And bProgramPress = False Then
    bProgramPress = True
    tabWizard.Tab = PreviousTab
  Else
    bProgramPress = False
  End If
  bButtonPress = False
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmSQLWizard.tabWizard_Click"
End Sub

Public Sub Initialise()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.Initialise()", etFullDebug

Dim objDatabase As pgDatabase
Dim objTable As pgTable
Dim objItem As ListItem
Dim rs As New Recordset

  tabWizard.Tab = 0
  cmdPrevious.Enabled = False
  
  StartMsg "Examining Server..."
  For Each objDatabase In svr.Databases
    If ((Not objDatabase.SystemObject) And (objDatabase.Status <> statInaccessible)) Then
      Set objItem = lvDatabases.ListItems.Add(, , objDatabase.Identifier, "database", "database")
      objItem.SubItems(1) = "No"
      For Each objTable In objDatabase.Tables
        If objTable.Identifier = "msysconf" Then
          objItem.SubItems(1) = "Yes"
          Set rs = objDatabase.Execute("SELECT nvalue FROM msysconf WHERE config = 101")
          If Not rs.EOF Then
            If rs!nvalue & "" = "1" Then
              objItem.SubItems(2) = "Yes"
            Else
              objItem.SubItems(2) = "No"
            End If
          Else
            objItem.SubItems(2) = "Not Set"
          End If
          Set rs = objDatabase.Execute("SELECT nvalue FROM msysconf WHERE config = 102")
          If Not rs.EOF Then
            objItem.SubItems(3) = rs!nvalue & ""
          Else
            objItem.SubItems(3) = "Not Set"
          End If
          Set rs = objDatabase.Execute("SELECT nvalue FROM msysconf WHERE config = 103")
          If Not rs.EOF Then
            objItem.SubItems(4) = rs!nvalue & ""
          Else
            objItem.SubItems(4) = "Not Set"
          End If
          Exit For
        End If
      Next objTable
    End If
  Next objDatabase
  EndMsg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Form_Load"
End Sub
