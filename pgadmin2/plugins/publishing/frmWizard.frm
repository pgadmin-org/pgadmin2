VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Publishing Wizard"
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
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":08CA
            Key             =   "database"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":0A24
            Key             =   "view"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":10F6
            Key             =   "function"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":1690
            Key             =   "index"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":1C2A
            Key             =   "language"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":22FC
            Key             =   "rule"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":29CE
            Key             =   "sequence"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":30A0
            Key             =   "table"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":3772
            Key             =   "trigger"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":3E44
            Key             =   "type"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":4516
            Key             =   "aggregate"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":4BE8
            Key             =   "operator"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":52BA
            Key             =   "server"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":5414
            Key             =   "domain"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picStrip 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      Picture         =   "frmWizard.frx":5AE6
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   19
      Top             =   0
      Width           =   465
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   330
      Left            =   5445
      TabIndex        =   3
      ToolTipText     =   "Move back a stage"
      Top             =   3960
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   6480
      TabIndex        =   2
      ToolTipText     =   "Return SQL and exit."
      Top             =   3960
      Visible         =   0   'False
      Width           =   960
   End
   Begin TabDlg.SSTab tabWizard 
      Height          =   3840
      Left            =   495
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   6773
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   176
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmWizard.frx":6BB5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblInfo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvDatabases"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmWizard.frx":6BD1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblInfo(1)"
      Tab(1).Control(1)=   "lvObjects"
      Tab(1).Control(2)=   "cmdObjectNone"
      Tab(1).Control(3)=   "cmdObjectAll"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   " "
      TabPicture(2)   =   "frmWizard.frx":6BED
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1"
      Tab(2).Control(1)=   "lblInfo(2)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   " "
      TabPicture(3)   =   "frmWizard.frx":6C09
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblInfo(4)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Picture2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   " "
      TabPicture(4)   =   "frmWizard.frx":6C25
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label1(3)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label1(2)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label1(1)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label1(0)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "lblInfo(3)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "lvServers"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "cmdAdd"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "cmdRemove"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "txtUsername"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "txtPassword"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "txtDatabase"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "txtPort"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "txtHost"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).ControlCount=   14
      TabCaption(5)   =   " "
      TabPicture(5)   =   "frmWizard.frx":6C41
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraStatus"
      Tab(5).Control(1)=   "lblInfo(5)"
      Tab(5).ControlCount=   2
      Begin VB.Frame fraStatus 
         Caption         =   "Status"
         Height          =   2535
         Left            =   -74820
         TabIndex        =   34
         Top             =   1125
         Width           =   6675
         Begin VB.Label lblOperation 
            Caption         =   "None."
            Height          =   195
            Left            =   585
            TabIndex        =   38
            Top             =   1845
            Width           =   5415
         End
         Begin VB.Label lblServer 
            Caption         =   "None."
            Height          =   195
            Left            =   585
            TabIndex        =   37
            Top             =   900
            Width           =   5415
         End
         Begin VB.Label Label2 
            Caption         =   "Current operation:"
            Height          =   240
            Index           =   1
            Left            =   225
            TabIndex        =   36
            Top             =   1395
            Width           =   1680
         End
         Begin VB.Label Label2 
            Caption         =   "Production server:"
            Height          =   240
            Index           =   0
            Left            =   225
            TabIndex        =   35
            Top             =   495
            Width           =   1680
         End
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   -74820
         TabIndex        =   12
         ToolTipText     =   "Enter the Hostname or IP Address of a production server."
         Top             =   2790
         Width           =   2445
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   -72300
         TabIndex        =   13
         Text            =   "5432"
         ToolTipText     =   "Enter the por on which the production server is listening."
         Top             =   2790
         Width           =   690
      End
      Begin VB.TextBox txtDatabase 
         Height          =   285
         Left            =   -71535
         TabIndex        =   14
         ToolTipText     =   "Enter the database to use on the production server."
         Top             =   2790
         Width           =   1860
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -72300
         PasswordChar    =   "*"
         TabIndex        =   16
         ToolTipText     =   "Enter the password to use on the production server."
         Top             =   3330
         Width           =   2625
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   -74820
         TabIndex        =   15
         ToolTipText     =   "Enter the username to use on the production server."
         Top             =   3330
         Width           =   2445
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   330
         Left            =   -69555
         TabIndex        =   18
         ToolTipText     =   "Remove the selected production server."
         Top             =   2745
         Width           =   1365
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   330
         Left            =   -69555
         TabIndex        =   17
         ToolTipText     =   "Add the production server details."
         Top             =   3195
         Width           =   1365
      End
      Begin VB.CommandButton cmdObjectAll 
         Height          =   555
         Left            =   -68655
         Picture         =   "frmWizard.frx":6C5D
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Select all objects."
         Top             =   1800
         Width           =   555
      End
      Begin VB.CommandButton cmdObjectNone 
         Height          =   555
         Left            =   -68655
         Picture         =   "frmWizard.frx":7527
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Deselect all objects."
         Top             =   2475
         Width           =   555
      End
      Begin MSComctlLib.ListView lvDatabases 
         Height          =   2535
         Left            =   135
         TabIndex        =   1
         Top             =   1170
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "Database/Schema"
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
         TabIndex        =   4
         Top             =   1170
         Width           =   6135
         _ExtentX        =   10821
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
            Text            =   "Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Comments"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSComctlLib.ListView lvServers 
         Height          =   1365
         Left            =   -74865
         TabIndex        =   11
         Top             =   1125
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   2408
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Hostname"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Port"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Database"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Username"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2580
         Left            =   -73425
         ScaleHeight     =   2580
         ScaleWidth      =   3300
         TabIndex        =   32
         Top             =   1035
         Width           =   3300
         Begin VB.OptionButton optDrop 
            Caption         =   "&Yes, DROP before CREATE"
            Height          =   195
            Index           =   0
            Left            =   450
            TabIndex        =   7
            Top             =   495
            Value           =   -1  'True
            Width           =   2670
         End
         Begin VB.OptionButton optDrop 
            Caption         =   "&No, just CREATE"
            Height          =   195
            Index           =   1
            Left            =   450
            TabIndex        =   8
            Top             =   1260
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2805
         Left            =   -73020
         ScaleHeight     =   2805
         ScaleWidth      =   3345
         TabIndex        =   33
         Top             =   765
         Width           =   3345
         Begin VB.OptionButton optReset 
            Caption         =   "&No, leave the sequence values alone."
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   10
            Top             =   1530
            Width           =   3120
         End
         Begin VB.OptionButton optReset 
            Caption         =   "&Yes, reset sequence values."
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   9
            Top             =   765
            Value           =   -1  'True
            Width           =   2670
         End
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmWizard.frx":7DF1
         Height          =   825
         Index           =   5
         Left            =   -74820
         TabIndex        =   31
         Top             =   180
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   "Do you want to reset the current value of any sequences that are published?"
         Height          =   735
         Index           =   4
         Left            =   -74820
         TabIndex        =   30
         Top             =   180
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmWizard.frx":7E9E
         Height          =   825
         Index           =   3
         Left            =   -74820
         TabIndex        =   29
         Top             =   180
         Width           =   6630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hostname/IP Address"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   28
         Top             =   2565
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Port"
         Height          =   195
         Index           =   1
         Left            =   -72300
         TabIndex        =   27
         Top             =   2565
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Database"
         Height          =   195
         Index           =   2
         Left            =   -71535
         TabIndex        =   26
         Top             =   2565
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         Height          =   195
         Index           =   3
         Left            =   -74820
         TabIndex        =   25
         Top             =   3105
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Index           =   4
         Left            =   -72300
         TabIndex        =   24
         Top             =   3105
         Width           =   690
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmWizard.frx":7FF2
         Height          =   735
         Index           =   2
         Left            =   -74820
         TabIndex        =   23
         Top             =   180
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   "Select the objects that you wish to publish."
         Height          =   735
         Index           =   1
         Left            =   -74820
         TabIndex        =   22
         Top             =   180
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   "Select the staging database/schema that you wish to publish to the remote server(s)."
         Height          =   735
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   180
         Width           =   6630
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   330
      Left            =   6480
      TabIndex        =   20
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

Private Sub cmdNext_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdNext_Click()", etFullDebug

Dim objItem As ListItem

  bButtonPress = True
  Select Case tabWizard.Tab
    Case 0
      'Only move on if at least one database is selected.
      If Not (lvDatabases.SelectedItem Is Nothing) Then
        txtDatabase.Text = fmtID(lvDatabases.SelectedItem.Tag.Database)
        GetObjects
        tabWizard.Tab = 1
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
      End If
    Case 1
      For Each objItem In lvObjects.ListItems
        If objItem.Checked Then
          tabWizard.Tab = 2
          cmdNext.Enabled = True
          cmdPrevious.Enabled = True
        End If
      Next objItem
      If tabWizard.Tab = 1 Then
        MsgBox "You must select at least one object to publish.", vbExclamation, "Error"
        lvObjects.SetFocus
        Exit Sub
      End If
    Case 2
      tabWizard.Tab = 3
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 3
      tabWizard.Tab = 4
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 4
      If lvServers.ListItems.Count > 0 Then
        tabWizard.Tab = 5
        cmdNext.Enabled = False
        cmdNext.Visible = False
        cmdOK.Enabled = True
        cmdOK.Visible = True
        cmdPrevious.Enabled = True
      Else
        MsgBox "You must add at least one production server.", vbExclamation, "Error"
        txtHost.SetFocus
      End If
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdNext_Click"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdAdd_Click()", etFullDebug

Dim objItem As ListItem

  If txtHost.Text = "" Then
    MsgBox "You must enter a hostname for the production server!", vbExclamation, "Error"
    txtHost.SetFocus
    Exit Sub
  End If
  If Val(txtPort.Text) < 1 Then
    MsgBox "You must enter a valid port number for the production server!", vbExclamation, "Error"
    txtPort.SetFocus
    Exit Sub
  End If
  If txtDatabase.Text = "" Then
    MsgBox "You must enter a database on the production server!", vbExclamation, "Error"
    txtDatabase.SetFocus
    Exit Sub
  End If
  If txtUsername.Text = "" Then
    MsgBox "You must enter a username for the production server!", vbExclamation, "Error"
    txtUsername.SetFocus
    Exit Sub
  End If

  Set objItem = lvServers.ListItems.Add(, , txtHost.Text, "server", "server")
  objItem.SubItems(1) = Val(txtPort.Text)
  objItem.SubItems(2) = txtDatabase.Text
  objItem.SubItems(3) = txtUsername.Text
  objItem.Tag = txtPassword.Text

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdAdd_Click"
End Sub

Private Sub cmdRemove_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdRemove_Click()", etFullDebug

  If Not (lvServers.SelectedItem Is Nothing) Then
    lvServers.ListItems.Remove lvServers.SelectedItem.Index
  Else
    MsgBox "You must select a server to remove!", vbExclamation, "Error"
    lvServers.SetFocus
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdRemove_Click"
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

Dim vObject As Variant
Dim vChildObject As Variant
Dim arrObjects() As Variant
Dim X As Long
Dim Y As Long
Dim vTemp As Variant
Dim objItem As ListItem
Dim szArgs As String
Dim vArg As Variant
Dim szSQL As String
Dim szROT As String
Dim szLOT As String
Dim cnPublish As New ADODB.Connection

  StartMsg "Publishing database..."

  'We will output the Schema object by object in OID order. This should work
  'most of the time unless (for example) a table is altered to use a later
  'created function as a default. Hopefully future releases of PostgreSQL
  'will include a pg_dependency table that we can use instead.
  
  'First, copy all objects into a single array: Aggregates, Functions,
  'Indexes, Operators, Rules, Sequences, Tables, Triggers,
  'Types & Views
  
  ReDim arrObjects(0)
  
  'Aggregates
  For Each vObject In lvDatabases.SelectedItem.Tag.Aggregates
    If Not vObject.SystemObject Then
      If lvObjects.ListItems("O" & vObject.OID).Checked Then
        Set arrObjects(UBound(arrObjects)) = vObject
        ReDim Preserve arrObjects(UBound(arrObjects) + 1)
      End If
    End If
  Next vObject
  
  'Domains
  For Each vObject In lvDatabases.SelectedItem.Tag.Domains
    If Not vObject.SystemObject Then
      If lvObjects.ListItems("O" & vObject.OID).Checked Then
        Set arrObjects(UBound(arrObjects)) = vObject
        ReDim Preserve arrObjects(UBound(arrObjects) + 1)
      End If
    End If
  Next vObject

  'Functions
  For Each vObject In lvDatabases.SelectedItem.Tag.Functions
    If Not vObject.SystemObject Then
      If lvObjects.ListItems("O" & vObject.OID).Checked Then
        Set arrObjects(UBound(arrObjects)) = vObject
        ReDim Preserve arrObjects(UBound(arrObjects) + 1)
      End If
    End If
  Next vObject

  'Operators
  For Each vObject In lvDatabases.SelectedItem.Tag.Operators
    If Not vObject.SystemObject Then
      If lvObjects.ListItems("O" & vObject.OID).Checked Then
        Set arrObjects(UBound(arrObjects)) = vObject
        ReDim Preserve arrObjects(UBound(arrObjects) + 1)
      End If
    End If
  Next vObject

  'Sequences
  For Each vObject In lvDatabases.SelectedItem.Tag.Sequences
    If Not vObject.SystemObject Then
      If lvObjects.ListItems("O" & vObject.OID).Checked Then
        Set arrObjects(UBound(arrObjects)) = vObject
        ReDim Preserve arrObjects(UBound(arrObjects) + 1)
      End If
    End If
  Next vObject

  'Tables
  For Each vObject In lvDatabases.SelectedItem.Tag.Tables
    If Not vObject.SystemObject Then
      If lvObjects.ListItems("O" & vObject.OID).Checked Then
        Set arrObjects(UBound(arrObjects)) = vObject
        ReDim Preserve arrObjects(UBound(arrObjects) + 1)
    
        'Indexes
        For Each vChildObject In vObject.Indexes
          If Not vObject.SystemObject Then
            If lvObjects.ListItems("O" & vObject.OID).Checked Then
              Set arrObjects(UBound(arrObjects)) = vChildObject
              ReDim Preserve arrObjects(UBound(arrObjects) + 1)
            End If
          End If
        Next vChildObject
      
        'Rules
        For Each vChildObject In vObject.Rules
          If Not vObject.SystemObject Then
            If lvObjects.ListItems("O" & vObject.OID).Checked Then
              Set arrObjects(UBound(arrObjects)) = vChildObject
              ReDim Preserve arrObjects(UBound(arrObjects) + 1)
            End If
          End If
        Next vChildObject
        
        'Triggers
        For Each vChildObject In vObject.Triggers
          If Not vObject.SystemObject Then
            If lvObjects.ListItems("O" & vObject.OID).Checked Then
              Set arrObjects(UBound(arrObjects)) = vChildObject
              ReDim Preserve arrObjects(UBound(arrObjects) + 1)
            End If
          End If
        Next vChildObject
      End If
    End If
  Next vObject
 
  'Types
  For Each vObject In lvDatabases.SelectedItem.Tag.Types
    If Not vObject.SystemObject Then
      If lvObjects.ListItems("O" & vObject.OID).Checked Then
        Set arrObjects(UBound(arrObjects)) = vObject
        ReDim Preserve arrObjects(UBound(arrObjects) + 1)
      End If
    End If
  Next vObject

  'Views
  For Each vObject In lvDatabases.SelectedItem.Tag.Views
    If Not vObject.SystemObject Then
      If lvObjects.ListItems("O" & vObject.OID).Checked Then
        Set arrObjects(UBound(arrObjects)) = vObject
        ReDim Preserve arrObjects(UBound(arrObjects) + 1)
      End If
    End If
  Next vObject
  
  'Lose the last empty element
  If UBound(arrObjects) > 0 Then ReDim Preserve arrObjects(UBound(arrObjects) - 1)
  
  'Now bubble sort the array by OID.
  For X = UBound(arrObjects) To LBound(arrObjects) Step -1
    For Y = LBound(arrObjects) + 1 To X
      If arrObjects(Y - 1).OID > arrObjects(Y).OID Then
        Set vTemp = arrObjects(Y - 1)
        Set arrObjects(Y - 1) = arrObjects(Y)
        Set arrObjects(Y) = vTemp
      End If
    Next Y
  Next X

  'Loop through the servers
  For Each objItem In lvServers.ListItems
    lblOperation.Caption = "Opening Connection: " & "DRIVER=PostgreSQL;SERVER=" & objItem.Text & ";PORT=" & objItem.SubItems(1) & ";UID=" & objItem.SubItems(3) & ";PWD=" & objItem.Tag & ";DATABASE=" & objItem.SubItems(2)
    lblOperation.Refresh
    svr.LogEvent "Opening Connection: " & "DRIVER=PostgreSQL;SERVER=" & objItem.Text & ";PORT=" & objItem.SubItems(1) & ";UID=" & objItem.SubItems(3) & ";PWD=" & objItem.Tag & ";DATABASE=" & objItem.SubItems(2), etSQL
    cnPublish.Open "DRIVER=PostgreSQL;SERVER=" & objItem.Text & ";PORT=" & objItem.SubItems(1) & ";UID=" & objItem.SubItems(3) & ";PWD=" & objItem.Tag & ";DATABASE=" & objItem.SubItems(2)
    lblServer.Caption = objItem.SubItems(2) & " on " & objItem.Text & ":" & objItem.SubItems(1)
    lblServer.Refresh
    
    'Loop through the object array...
    For X = 0 To UBound(arrObjects)
      'Drop the object first if required.
      If optDrop(0).Value Then
        lblOperation.Caption = "Dropping " & arrObjects(X).ObjectType & ": " & arrObjects(X).Identifier
        lblOperation.Refresh
        Select Case arrObjects(X).ObjectType
          Case "Aggregate"
            szSQL = "DROP AGGREGATE " & arrObjects(X).FormattedID & " *"
            
          Case "Domain"
            szSQL = "DROP DOMAIN " & arrObjects(X).FormattedID
 
          Case "Function"
            szSQL = "DROP FUNCTION " & arrObjects(X).FormattedID
 
          Case "Operator"
            szSQL = "DROP OPERATOR " & arrObjects(X).FormattedID
 
          Case "Sequence"
            szSQL = "DROP SEQUENCE " & arrObjects(X).FormattedID
            
          Case "Table"
            szSQL = "DROP TABLE " & arrObjects(X).FormattedID
          
          Case "Index"
            szSQL = "DROP INDEX " & arrObjects(X).FormattedID
            
          Case "Rule"
            szSQL = "DROP RULE " & arrObjects(X).FormattedID
            
          Case "Trigger"
            szSQL = "DROP TRIGGER " & arrObjects(X).FormattedID
            
          Case "Type"
            szSQL = "DROP TYPE " & arrObjects(X).FormattedID
            
          Case "View"
            szSQL = "DROP VIEW " & arrObjects(X).FormattedID
            
        End Select
        svr.LogEvent "SQL (" & objItem.SubItems(2) & " on " & objItem.Text & ":" & objItem.SubItems(1) & "): " & szSQL, etSQL
        On Error Resume Next
        cnPublish.Execute szSQL
        On Error GoTo Err_Handler
      End If
      
      'Now Create the new object
      lblOperation.Caption = "Creating " & arrObjects(X).ObjectType & ": " & arrObjects(X).Identifier
      lblOperation.Refresh
      If arrObjects(X).ObjectType = "Sequence" Then
        svr.LogEvent "SQL (" & objItem.SubItems(2) & " on " & objItem.Text & ":" & objItem.SubItems(1) & "): " & arrObjects(X).SQL(optReset(0).Value), etSQL
        cnPublish.Execute arrObjects(X).SQL(optReset(0).Value)
      Else
        svr.LogEvent "SQL (" & objItem.SubItems(2) & " on " & objItem.Text & ":" & objItem.SubItems(1) & "): " & arrObjects(X).SQL, etSQL
        cnPublish.Execute arrObjects(X).SQL
      End If
      
    Next X
  
    lblOperation.Caption = "Closing Connection: " & "DRIVER=PostgreSQL;SERVER=" & objItem.Text & ";PORT=" & objItem.SubItems(1) & ";UID=" & objItem.SubItems(3) & ";PWD=" & objItem.Tag & ";DATABASE=" & objItem.SubItems(2)
    If cnPublish.State <> adStateClosed Then cnPublish.Close
  Next objItem

  EndMsg
  bRunning = False
  Unload Me
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdOK_Click"
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdPrevious_Click()", etFullDebug

  bButtonPress = True
  Select Case tabWizard.Tab
    Case 5
      tabWizard.Tab = 4
      cmdNext.Enabled = True
      cmdNext.Visible = True
      cmdOK.Enabled = False
      cmdOK.Visible = False
      cmdPrevious.Enabled = True
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

Private Sub cmdObjectAll_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdObjectAll_Click()", etFullDebug

Dim objItem As ListItem

  For Each objItem In lvObjects.ListItems
    objItem.Checked = True
  Next objItem
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdObjectAll_Click"
End Sub

Private Sub cmdObjectNone_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdObjectNone_Click()", etFullDebug

Dim objItem As ListItem

  For Each objItem In lvObjects.ListItems
    objItem.Checked = False
  Next objItem
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdObjectNone_Click"
End Sub

Public Sub Initialise()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.Initialise()", etFullDebug

Dim objDatabase As pgDatabase
Dim objNamespace As pgNamespace
Dim objItem As ListItem
  
  lvDatabases.ListItems.Clear
  tabWizard.Tab = 0
  cmdPrevious.Enabled = False
  
  StartMsg "Examining Server..."
  If svr.dbVersion.VersionNum >= 7.3 Then
    For Each objDatabase In svr.Databases
      If ((Not objDatabase.SystemObject) And (Not (objDatabase.Namespaces Is Nothing))) Then
        For Each objNamespace In objDatabase.Namespaces
          If (Not objNamespace.SystemObject) Or (objNamespace.Name = "public") Then
            Set objItem = lvDatabases.ListItems.Add(, , objDatabase.FormattedID & "." & objNamespace.FormattedID, "database", "database")
            objItem.SubItems(1) = Replace(objNamespace.Comment, vbCrLf, " ")
            Set objItem.Tag = objNamespace
          End If
        Next objNamespace
      End If
    Next objDatabase
  Else
    For Each objDatabase In svr.Databases
      If ((Not objDatabase.SystemObject) And (Not (objDatabase.Namespaces Is Nothing))) Then
        Set objItem = lvDatabases.ListItems.Add(, , objDatabase.Identifier, "database", "database")
        objItem.SubItems(1) = Replace(objDatabase.Comment, vbCrLf, " ")
        Set objItem.Tag = objDatabase.Namespaces("public")
      End If
    Next objDatabase
  End If
  
  txtUsername.Text = svr.Username
  txtPort.Text = svr.Port
  txtHost.Text = svr.Server
  
  EndMsg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Initialise"
End Sub

Private Sub GetObjects()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.GetObjects()", etFullDebug

Dim objItem As ListItem
Dim vObject As Variant
Dim vChildObject As Variant

  lvObjects.ListItems.Clear
  StartMsg "Examining Server..."
  
  'Aggregates
  For Each vObject In lvDatabases.SelectedItem.Tag.Aggregates
    If Not vObject.SystemObject Then
      Set objItem = lvObjects.ListItems.Add(, "O" & vObject.OID, vObject.FormattedID, "aggregate", "aggregate")
      objItem.SubItems(1) = "Aggregate"
      objItem.SubItems(2) = Replace(vObject.Comment, vbCrLf, " ")
    End If
  Next vObject
  
  'Domains
  For Each vObject In lvDatabases.SelectedItem.Tag.Domains
    If Not vObject.SystemObject Then
      Set objItem = lvObjects.ListItems.Add(, "O" & vObject.OID, vObject.FormattedID, "domain", "domain")
      objItem.SubItems(1) = "Domain"
      objItem.SubItems(2) = Replace(vObject.Comment, vbCrLf, " ")
    End If
  Next vObject

  'Functions
  For Each vObject In lvDatabases.SelectedItem.Tag.Functions
    If Not vObject.SystemObject Then
      Set objItem = lvObjects.ListItems.Add(, "O" & vObject.OID, vObject.FormattedID, "function", "function")
      objItem.SubItems(1) = "Function"
      objItem.SubItems(2) = Replace(vObject.Comment, vbCrLf, " ")
    End If
  Next vObject

  'Operators
  For Each vObject In lvDatabases.SelectedItem.Tag.Operators
    If Not vObject.SystemObject Then
      Set objItem = lvObjects.ListItems.Add(, "O" & vObject.OID, vObject.FormattedID, "operator", "operator")
      objItem.SubItems(1) = "Operator"
      objItem.SubItems(2) = Replace(vObject.Comment, vbCrLf, " ")
    End If
  Next vObject

  'Sequences
  For Each vObject In lvDatabases.SelectedItem.Tag.Sequences
    If Not vObject.SystemObject Then
      Set objItem = lvObjects.ListItems.Add(, "O" & vObject.OID, vObject.FormattedID, "sequence", "sequence")
      objItem.SubItems(1) = "Sequence"
      objItem.SubItems(2) = Replace(vObject.Comment, vbCrLf, " ")
    End If
  Next vObject

  'Tables
  For Each vObject In lvDatabases.SelectedItem.Tag.Tables
    If Not vObject.SystemObject Then
      Set objItem = lvObjects.ListItems.Add(, "O" & vObject.OID, vObject.FormattedID, "table", "table")
      objItem.SubItems(1) = "Table"
      objItem.SubItems(2) = Replace(vObject.Comment, vbCrLf, " ")
    
      'Indexes
      For Each vChildObject In vObject.Indexes
        If Not vChildObject.SystemObject Then
          Set objItem = lvObjects.ListItems.Add(, "O" & vChildObject.OID, vChildObject.FormattedID, "index", "index")
          objItem.SubItems(1) = "Index"
          objItem.SubItems(2) = Replace(vChildObject.Comment, vbCrLf, " ")
        End If
      Next vChildObject
      
      'Rules
      For Each vChildObject In vObject.Rules
        If Not vChildObject.SystemObject Then
          Set objItem = lvObjects.ListItems.Add(, "O" & vChildObject.OID, vChildObject.FormattedID, "rule", "rule")
          objItem.SubItems(1) = "Rule"
          objItem.SubItems(2) = Replace(vChildObject.Comment, vbCrLf, " ")
        End If
      Next vChildObject
      
      'Triggers
      For Each vChildObject In vObject.Triggers
        If Not vChildObject.SystemObject Then
          Set objItem = lvObjects.ListItems.Add(, "O" & vChildObject.OID, vChildObject.FormattedID, "trigger", "trigger")
          objItem.SubItems(1) = "Trigger"
          objItem.SubItems(2) = Replace(vChildObject.Comment, vbCrLf, " ")
        End If
      Next vChildObject
    End If
  Next vObject

  'Types
  For Each vObject In lvDatabases.SelectedItem.Tag.Types
    If Not vObject.SystemObject Then
      Set objItem = lvObjects.ListItems.Add(, "O" & vObject.OID, vObject.FormattedID, "type", "type")
      objItem.SubItems(1) = "Type"
      objItem.SubItems(2) = Replace(vObject.Comment, vbCrLf, " ")
    End If
  Next vObject

  'Views
  For Each vObject In lvDatabases.SelectedItem.Tag.Views
    If Not vObject.SystemObject Then
      Set objItem = lvObjects.ListItems.Add(, "O" & vObject.OID, vObject.FormattedID, "view", "view")
      objItem.SubItems(1) = "View"
      objItem.SubItems(2) = Replace(vObject.Comment, vbCrLf, " ")
    End If
  Next vObject
  EndMsg
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.GetObjects"
End Sub

