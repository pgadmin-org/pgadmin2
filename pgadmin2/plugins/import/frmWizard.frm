VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Wizard"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7530
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   1170
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":08CA
            Key             =   "database"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":0A24
            Key             =   "table"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":0B7E
            Key             =   "column"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picStrip 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      Picture         =   "frmWizard.frx":1118
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   4
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
      ToolTipText     =   "Run the Data Import"
      Top             =   3960
      Visible         =   0   'False
      Width           =   960
   End
   Begin TabDlg.SSTab tabWizard 
      Height          =   3840
      Left            =   495
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Check if the Import Wizard should expect to find a trailing delimiter character at the end of each record in the import file."
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
      TabPicture(0)   =   "frmWizard.frx":1D55
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblInfo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvDatabases"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmWizard.frx":1D71
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblInfo(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lvTables"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " "
      TabPicture(2)   =   "frmWizard.frx":1D8D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblInfo(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lvColumns"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdColumnUp"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdColumnDown"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdColumnAll"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdColumnNone"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   " "
      TabPicture(3)   =   "frmWizard.frx":1DA9
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblInfo(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label1"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "txtSample"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "txtFile"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdBrowse"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      TabCaption(4)   =   " "
      TabPicture(4)   =   "frmWizard.frx":1DC5
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblInfo(6)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label3"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label4"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label5"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "txtDelimiter"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "txtQuote"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "txtAsciiDelimiter"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "txtAsciiQuote"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "chkTrailing"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).ControlCount=   10
      TabCaption(5)   =   " "
      TabPicture(5)   =   "frmWizard.frx":1DE1
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblInfo(7)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "lvSubstitutions"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "txtSubFind"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "txtSubReplace"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "cmdSubAdd"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "cmdSubRemove"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).ControlCount=   6
      TabCaption(6)   =   " "
      TabPicture(6)   =   "frmWizard.frx":1DFD
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblInfo(4)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "picStatus"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).ControlCount=   2
      Begin VB.PictureBox picStatus 
         Height          =   2445
         Left            =   -74865
         ScaleHeight     =   2385
         ScaleWidth      =   6660
         TabIndex        =   37
         Top             =   1260
         Width           =   6720
         Begin MSComctlLib.ProgressBar pbStatus 
            Height          =   285
            Left            =   225
            TabIndex        =   38
            Top             =   1710
            Width           =   6225
            _ExtentX        =   10980
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label Label8 
            Caption         =   "Progress"
            Height          =   240
            Left            =   225
            TabIndex        =   43
            Top             =   1485
            Width           =   1590
         End
         Begin VB.Label lblErrors 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   240
            Left            =   3195
            TabIndex        =   42
            Top             =   990
            Width           =   1590
         End
         Begin VB.Label lblRecords 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   240
            Left            =   3195
            TabIndex        =   41
            Top             =   495
            Width           =   1590
         End
         Begin VB.Label Label7 
            Caption         =   "Errors encountered:"
            Height          =   240
            Left            =   1440
            TabIndex        =   40
            Top             =   990
            Width           =   1590
         End
         Begin VB.Label Label6 
            Caption         =   "Records imported:"
            Height          =   240
            Left            =   1440
            TabIndex        =   39
            Top             =   495
            Width           =   1590
         End
      End
      Begin VB.CommandButton cmdSubRemove 
         Caption         =   "&Remove"
         Height          =   330
         Left            =   -69150
         TabIndex        =   36
         ToolTipText     =   "Remove the selected substitution."
         Top             =   2970
         Width           =   1005
      End
      Begin VB.CommandButton cmdSubAdd 
         Caption         =   "&Add"
         Height          =   330
         Left            =   -69150
         TabIndex        =   35
         ToolTipText     =   "Add the defined substitution."
         Top             =   3375
         Width           =   1005
      End
      Begin VB.TextBox txtSubReplace 
         Height          =   285
         Left            =   -72030
         TabIndex        =   34
         ToolTipText     =   "Enter the string to replace with."
         Top             =   3420
         Width           =   2760
      End
      Begin VB.TextBox txtSubFind 
         Height          =   285
         Left            =   -74865
         TabIndex        =   33
         ToolTipText     =   "Enter the string to search for."
         Top             =   3420
         Width           =   2760
      End
      Begin VB.CommandButton cmdColumnNone 
         Height          =   555
         Left            =   -68655
         Picture         =   "frmWizard.frx":1E19
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Deselect all columns"
         Top             =   2520
         Width           =   555
      End
      Begin VB.CommandButton cmdColumnAll 
         Height          =   555
         Left            =   -68655
         Picture         =   "frmWizard.frx":26E3
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Select all columns"
         Top             =   1890
         Width           =   555
      End
      Begin VB.CommandButton cmdColumnDown 
         Height          =   555
         Left            =   -68655
         Picture         =   "frmWizard.frx":2FAD
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Move the selected column down."
         Top             =   3150
         Width           =   555
      End
      Begin VB.CommandButton cmdColumnUp 
         Height          =   555
         Left            =   -68655
         Picture         =   "frmWizard.frx":3877
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Move the selected column up."
         Top             =   1260
         Width           =   555
      End
      Begin VB.CheckBox chkTrailing 
         Caption         =   "Expect a trailing delimiter character?"
         Height          =   240
         Left            =   -72930
         TabIndex        =   27
         Top             =   2160
         Width           =   3120
      End
      Begin VB.TextBox txtAsciiQuote 
         Height          =   285
         Left            =   -70635
         MaxLength       =   3
         TabIndex        =   24
         ToolTipText     =   "Enter the ASCII value for the quote."
         Top             =   2925
         Width           =   825
      End
      Begin VB.TextBox txtAsciiDelimiter 
         Height          =   285
         Left            =   -70635
         MaxLength       =   3
         TabIndex        =   23
         ToolTipText     =   "Enter the ASCII value for the delimiter."
         Top             =   1755
         Width           =   825
      End
      Begin VB.TextBox txtQuote 
         Height          =   285
         Left            =   -72930
         MaxLength       =   1
         TabIndex        =   22
         ToolTipText     =   "Enter a quote character."
         Top             =   2925
         Width           =   825
      End
      Begin VB.TextBox txtDelimiter 
         Height          =   285
         Left            =   -72930
         MaxLength       =   1
         TabIndex        =   21
         ToolTipText     =   "Enter a delimiter character."
         Top             =   1755
         Width           =   825
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   -68520
         TabIndex        =   18
         ToolTipText     =   "Browse for the Import file."
         Top             =   1215
         Width           =   330
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   -74100
         TabIndex        =   17
         ToolTipText     =   "Enter (or select) a filename to import data from."
         Top             =   1215
         Width           =   5550
      End
      Begin VB.TextBox txtSample 
         BackColor       =   &H8000000F&
         Height          =   2130
         Left            =   -74865
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   15
         Top             =   1575
         Width           =   6675
      End
      Begin MSComctlLib.ListView lvDatabases 
         Height          =   2445
         Left            =   135
         TabIndex        =   0
         Top             =   1260
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   4313
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
            Text            =   "Database"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Comments"
            Object.Width           =   7498
         EndProperty
      End
      Begin MSComctlLib.ListView lvTables 
         Height          =   2445
         Left            =   -74865
         TabIndex        =   13
         Top             =   1260
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   4313
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
            Text            =   "Database"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Comments"
            Object.Width           =   7498
         EndProperty
      End
      Begin MSComctlLib.ListView lvColumns 
         Height          =   2445
         Left            =   -74865
         TabIndex        =   14
         Top             =   1260
         Width           =   6180
         _ExtentX        =   10901
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Database"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Comments"
            Object.Width           =   7498
         EndProperty
      End
      Begin MSComctlLib.ListView lvSubstitutions 
         Height          =   2040
         Left            =   -74865
         TabIndex        =   32
         Top             =   1260
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3598
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Search for"
            Object.Width           =   4868
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Replace with"
            Object.Width           =   4868
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "ASCII"
         Height          =   240
         Left            =   -71310
         TabIndex        =   26
         Top             =   2970
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "ASCII"
         Height          =   240
         Left            =   -71310
         TabIndex        =   25
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label Label3 
         Caption         =   "Quote"
         Height          =   240
         Left            =   -73695
         TabIndex        =   20
         Top             =   2970
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "Delimiter"
         Height          =   240
         Left            =   -73695
         TabIndex        =   19
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Import file"
         Height          =   195
         Left            =   -74865
         TabIndex        =   16
         Top             =   1260
         Width           =   780
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmWizard.frx":4141
         Height          =   825
         Index           =   4
         Left            =   -74820
         TabIndex        =   12
         Top             =   360
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmWizard.frx":41D7
         Height          =   825
         Index           =   7
         Left            =   -74820
         TabIndex        =   11
         Top             =   360
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmWizard.frx":4310
         Height          =   825
         Index           =   6
         Left            =   -74820
         TabIndex        =   10
         Top             =   360
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmWizard.frx":4440
         Height          =   825
         Index           =   3
         Left            =   -74820
         TabIndex        =   9
         Top             =   360
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmWizard.frx":4581
         Height          =   825
         Index           =   2
         Left            =   -74820
         TabIndex        =   8
         Top             =   360
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   "Select the target table for the imported data."
         Height          =   735
         Index           =   1
         Left            =   -74820
         TabIndex        =   7
         Top             =   360
         Width           =   6630
      End
      Begin VB.Label lblInfo 
         Caption         =   "Select the database containing the target table for the imported data."
         Height          =   735
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   6630
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   330
      Left            =   6480
      TabIndex        =   5
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
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit
Dim bButtonPress As Boolean
Dim bProgramPress As Boolean
Dim bUpdating As Boolean

Private Sub chkTrailing_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.chkTrailing_Click()", etFullDebug

  If chkTrailing.Value = 1 Then
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\Import Wizard", "Trailing Delimiter", regString, "Y"
  Else
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\Import Wizard", "Trailing Delimiter", regString, "N"
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.chkTrailing_Click"
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdBrowse_Click()", etFullDebug

Dim szData As String
Dim X As Integer
Dim fNum As Integer

  cdlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
  cdlg.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
  cdlg.ShowOpen

  If cdlg.FileName = "" Then Exit Sub
  txtFile.Text = cdlg.FileName
  txtSample.Text = ""
  fNum = FreeFile
  Open txtFile.Text For Input As #fNum
  For X = 0 To 10
    If Not EOF(fNum) Then
      Line Input #fNum, szData
      txtSample.Text = txtSample.Text & szData & vbCrLf
    End If
  Next
  Close #fNum
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdBrowse_Click"
End Sub

Private Sub cmdColumnAll_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdColumnAll_Click()", etFullDebug

Dim objItem As ListItem

  For Each objItem In lvColumns.ListItems
    objItem.Checked = True
  Next objItem
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdColumnAll_Click"
End Sub

Private Sub cmdColumnNone_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdColumnNone_Click()", etFullDebug

Dim objItem As ListItem

  For Each objItem In lvColumns.ListItems
    objItem.Checked = False
  Next objItem
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdColumnNone_Click"
End Sub

Private Sub cmdColumnUp_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdColumnUp_Click()", etFullDebug

Dim lIndex As Long
Dim objItem As ListItem

  'Exit if we're moving the first item up.
  If lvColumns.SelectedItem.Index <= 1 Then Exit Sub
  
  'Remember the initial Index
  lIndex = lvColumns.SelectedItem.Index
  
  'Create a new item
  Set objItem = lvColumns.ListItems.Add(lIndex - 1, , lvColumns.SelectedItem.Text, lvColumns.SelectedItem.Icon, lvColumns.SelectedItem.SmallIcon)
  objItem.Checked = lvColumns.SelectedItem.Checked
  objItem.SubItems(1) = lvColumns.SelectedItem.SubItems(1)
  
  'Remove the old item and select the new.
  lvColumns.ListItems.Remove lvColumns.SelectedItem.Index
  objItem.EnsureVisible
  objItem.Selected = True
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdColumnUp_Click"
End Sub

Private Sub cmdColumnDown_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdColumnDown_Click()", etFullDebug

Dim lIndex As Long
Dim lMax As Long
Dim objItem As ListItem

  'Exit if we're moving the last item down.
  For Each objItem In lvColumns.ListItems
    If objItem.Index > lMax Then lMax = objItem.Index
  Next objItem
  If lvColumns.SelectedItem.Index >= lMax Then Exit Sub
  
  'Remember the initial Index
  lIndex = lvColumns.SelectedItem.Index
  
  'Create a new item
  Set objItem = lvColumns.ListItems.Add(lIndex + 2, , lvColumns.SelectedItem.Text, lvColumns.SelectedItem.Icon, lvColumns.SelectedItem.SmallIcon)
  objItem.Checked = lvColumns.SelectedItem.Checked
  objItem.SubItems(1) = lvColumns.SelectedItem.SubItems(1)
  
  'Remove the old item and select the new.
  lvColumns.ListItems.Remove lvColumns.SelectedItem.Index
  objItem.EnsureVisible
  objItem.Selected = True
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdColumnDown_Click"
End Sub

Private Sub cmdNext_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdNext_Click()", etFullDebug

Dim objItem As ListItem

  bButtonPress = True
  Select Case tabWizard.Tab
    Case 0
      'Only move on if a database is selected.
      If Not (lvDatabases.SelectedItem Is Nothing) Then
        GetTables
        tabWizard.Tab = 1
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
      Else
        MsgBox "You must select a database.", vbExclamation, "Error"
        lvDatabases.SetFocus
        Exit Sub
      End If
    Case 1
      'Only move on if a database is selected.
      If Not (lvTables.SelectedItem Is Nothing) Then
        GetColumns
        tabWizard.Tab = 2
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
      Else
        MsgBox "You must select a table.", vbExclamation, "Error"
        lvTables.SetFocus
        Exit Sub
      End If
    Case 2
      'Only move on if at least one column is selected.
      For Each objItem In lvColumns.ListItems
        If objItem.Checked Then
          tabWizard.Tab = 3
          cmdNext.Enabled = True
          cmdPrevious.Enabled = True
          Exit For
        End If
      Next objItem
      If tabWizard.Tab = 2 Then
        MsgBox "You must select at least one column.", vbExclamation, "Error"
        lvColumns.SetFocus
        Exit Sub
      End If
    Case 3
      'Only move on if the file exists
      If (txtFile.Text = "") Or (Dir(txtFile.Text) = "") Then
        MsgBox "An invalid filename was specified!", vbExclamation, "Error"
        txtFile.SetFocus
        Exit Sub
      End If
      GetFormatting
      tabWizard.Tab = 4
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 4
      'Only move on if at least a delimiter has been specified
      If txtDelimiter.Text = "" Then
        MsgBox "No delimiter was specified!", vbExclamation, "Error"
        txtDelimiter.SetFocus
        Exit Sub
      End If
      GetSubstitutions
      tabWizard.Tab = 5
      cmdNext.Enabled = True
      cmdPrevious.Enabled = True
    Case 5
      tabWizard.Tab = 6
      cmdNext.Enabled = False
      cmdNext.Visible = False
      cmdOK.Enabled = True
      cmdOK.Visible = True
      cmdPrevious.Enabled = True
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdNext_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdOK_Click()", etFullDebug

Dim fNum As Integer
Dim X As Long
Dim lTuple As Long
Dim lErrors As Long
Dim lCol As Long
Dim lColumns As Long
Dim objItem As ListItem
Dim szDatabase As String
Dim szTable As String
Dim szColumns As String
Dim szCols() As String
Dim szRows As String
Dim szData() As String
Dim szRawData As String
Dim szSQL As String
Dim szQuote As String
Dim szDelimiter As String
Dim szChar As String
Dim bSQlErrors As Boolean
Dim bTrailing As Boolean
Dim bHaveRow As Boolean
Dim bInQuote As Boolean

  StartMsg "Importing data..."
      
    'Open the import file
    If Dir(txtFile.Text) = "" Then
      MsgBox "The import file could not be found!", vbExclamation, "Error"
      tabWizard.Tab = 3
      txtFile.SetFocus
      Exit Sub
    End If
    pbStatus.Max = FileLen(txtFile.Text) + 2
    pbStatus.Value = 0
    fNum = FreeFile
    Open txtFile.Text For Input As #fNum
    
    'Dim the Data Array & build the column array
    For Each objItem In lvColumns.ListItems
      lColumns = lColumns + 1
    Next objItem
    ReDim szData(lColumns)
    ReDim szCols(lColumns)
    X = 1
    For Each objItem In lvColumns.ListItems
      szCols(X) = QUOTE & objItem.Text & QUOTE
      X = X + 1
    Next objItem
    
    'Store some values for fast access
    szDatabase = lvDatabases.SelectedItem.Text
    szTable = QUOTE & lvTables.SelectedItem.Text & QUOTE
    szQuote = txtQuote.Text
    szDelimiter = txtDelimiter.Text
    If chkTrailing.Value = 1 Then
      bTrailing = True
    Else
      bTrailing = False
    End If
    
    'Start a transaction
    svr.Databases(szDatabase).Execute "BEGIN"
    
    'Now we do the actual import...
    While Not EOF(fNum)
      lTuple = lTuple + 1
      
      'Loop through reading and processing lines until we have a complete row.
      lCol = 1
      While Not bHaveRow
        Line Input #fNum, szRawData
        pbStatus.Value = pbStatus.Value + Len(szRawData) + 2
        
        'If we're already in quotes then we must have just had a CRLF
        'We'll add vbCrLf as it will get convereted to \n later
        If bInQuote Then szData(lCol) = szData(lCol) & vbCrLf
        
        'Scan the data a char at a time
        For X = 1 To Len(szRawData)
          szChar = Mid(szRawData, X, 1)
          If szChar = szQuote Then
          
            'The current char is a quote - check to see if the next char is
            'as well.
            If Mid(szRawData, X + 1, 1) = szQuote Then
            
              'Yes the next char is a quote, so add a quote to the string and
              'move past the second quote.
              szData(lCol) = szData(lCol) & szQuote
              X = X + 1
            Else
            
              'There was no second quote, so that must have been an actual quote...
              If bInQuote Then
                bInQuote = False
              Else
                bInQuote = True
              End If
            End If
          
          'If the char is a delimiter then...
          ElseIf szChar = szDelimiter Then
            
            'If we're in quotes then the delimiter is part of a string...
            If bInQuote Then
              szData(lCol) = szData(lCol) & szChar
            Else
              lCol = lCol + 1
              If bTrailing Then
                If lCol = lColumns Then bHaveRow = True
              Else
                If lCol = lColumns - 1 Then bHaveRow = True
              End If
            End If
          Else
          
            'Ahh, a regular character...
            szData(lCol) = szData(lCol) & szChar
          End If
        Next X
      Wend
      
      'Now we have the complete row, build and execute the SQL.
      'We'll also run the data through the Substitutions here.
      For X = 1 To lColumns
        If Not ((szData(X) = "") Or (szData(X) = szQuote)) Then
          szColumns = szColumns & szCols(X) & ", "
          For Each objItem In lvSubstitutions.ListItems
            szData(X) = Replace(szData(X), objItem.Text, objItem.SubItems(1))
          Next objItem
          szRows = szRows & "'" & dbSZ(szData(X)) & "', "
        End If
      Next X
      If Len(szColumns) > 2 Then szColumns = Left(szColumns, Len(szColumns) - 2)
      If Len(szRows) > 2 Then szRows = Left(szRows, Len(szRows) - 2)
      
      szSQL = "INSERT INTO " & szTable & " (" & szColumns & ") VALUES (" & szRows & ")"
      bSQlErrors = True
      svr.Databases(szDatabase).Execute szSQL
Reset:
      bSQlErrors = False
      bHaveRow = False
      bInQuote = False
      For X = 1 To lColumns
        szData(X) = ""
      Next X
      szRows = ""
      szColumns = ""
      lblRecords.Caption = lTuple
      lblRecords.Refresh
    Wend
  EndMsg
  
  If lErrors > 0 Then
    MsgBox lErrors & " error(s) were encountered." & vbCrLf & vbCrLf & "The data import will not be committed.", vbExclamation, "Import Complete"
    svr.LogEvent lErrors & " error(s) were encountered." & vbCrLf & vbCrLf & "The data import will not be committed.", etMiniDebug
    svr.Databases(szDatabase).Execute "ROLLBACK"
    MsgBox "Import rolled back! Please check the import data and try again.", vbExclamation, "Data Import"
    Exit Sub
  Else
    svr.LogEvent "Comitting data import...", etMiniDebug
    svr.Databases(szDatabase).Execute "COMMIT"
    svr.LogEvent "The data import from " & txtFile.Text & " has completed." & vbCrLf & vbCrLf & "Records processed: " & lTuple & vbCrLf & "Errors encountered: " & lErrors, etMiniDebug
    MsgBox "The data import from " & txtFile.Text & " has completed." & vbCrLf & vbCrLf & "Records processed: " & lTuple & vbCrLf & "Errors encountered: " & lErrors, vbInformation, "Import Complete"
  End If
  
  bRunning = False
  Unload Me
  
  Exit Sub
Err_Handler:
  If bSQlErrors = True Then
    lErrors = lErrors + 1
    lblErrors.Caption = lErrors
    lblErrors.Refresh
    svr.LogEvent "An error occured importing row " & lTuple + 1 & ". The generated SQL was: " & szSQL, etErrors
    GoTo Reset
  End If
  If Err.Number <> 0 Then
    EndMsg
    LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdOK_Click"
  End If
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
      tabWizard.Tab = 4
      cmdNext.Enabled = True
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

Private Sub cmdSubRemove_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdSubRemove_Click()", etFullDebug

Dim X As Long

  If Not (lvSubstitutions.SelectedItem Is Nothing) Then
    If MsgBox("Are you sure you wish to delete the selected substitution?", vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then Exit Sub
    lvSubstitutions.ListItems.Remove lvSubstitutions.SelectedItem.Index
  End If
  
  'Update the Substitutions in the Registry
  RegDelSubkey HKEY_CURRENT_USER, "Software\pgAdmin II\" & App.Title & "\Substitution Map"
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\" & App.Title & "\Substitution Map", "Entries", regString, lvSubstitutions.ListItems.Count
  If lvSubstitutions.ListItems.Count > 0 Then
    For X = 1 To lvSubstitutions.ListItems.Count
      RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\" & App.Title & "\Substitution Map", "Search - " & X, regString, lvSubstitutions.ListItems(X).Text
      RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\" & App.Title & "\Substitution Map", "Replace - " & X, regString, lvSubstitutions.ListItems(X).SubItems(1)
    Next
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdSubRemove_Click"
End Sub

Private Sub cmdSubAdd_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdSubAdd_Click()", etFullDebug

Dim objItem As ListItem
Dim X As Long

  If txtSubFind.Text = "" Then
    MsgBox "You must enter a string to search for!", vbExclamation, "Error"
    txtSubFind.SetFocus
    Exit Sub
  End If
  Set objItem = lvSubstitutions.ListItems.Add(, , txtSubFind.Text)
  objItem.SubItems(1) = txtSubReplace.Text
  
  'Update the Substitutions in the Registry
  RegDelSubkey HKEY_CURRENT_USER, "Software\pgAdmin II\" & App.Title & "\Substitution Map"
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\" & App.Title & "\Substitution Map", "Entries", regString, lvSubstitutions.ListItems.Count
  If lvSubstitutions.ListItems.Count > 0 Then
    For X = 1 To lvSubstitutions.ListItems.Count
      RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\" & App.Title & "\Substitution Map", "Search - " & X, regString, lvSubstitutions.ListItems(X).Text
      RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\" & App.Title & "\Substitution Map", "Replace - " & X, regString, lvSubstitutions.ListItems(X).SubItems(1)
    Next
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdSubAdd_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.Form_Unload()", etFullDebug

  bRunning = False

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Form_Unload"
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

  tabWizard.Tab = 0
  cmdPrevious.Enabled = False
  
  GetDatabases
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Form_Load"
End Sub

Private Sub GetDatabases()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.GetDatabases()", etFullDebug

Dim objItem As ListItem
Dim objDatabase As pgDatabase

  lvDatabases.ListItems.Clear
  StartMsg "Examining Server..."
  For Each objDatabase In svr.Databases
    If Not objDatabase.SystemObject Then
      Set objItem = lvDatabases.ListItems.Add(, , objDatabase.Identifier, "database", "database")
      objItem.SubItems(1) = Replace(objDatabase.Comment, vbCrLf, " ")
    End If
  Next objDatabase
  EndMsg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.GetDatabases"
End Sub

Private Sub GetTables()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.GetTables()", etFullDebug

Dim objItem As ListItem
Dim objTable As pgTable

  lvTables.ListItems.Clear
  StartMsg "Examining Server..."
  For Each objTable In svr.Databases(lvDatabases.SelectedItem.Text).Tables
    If Not objTable.SystemObject Then
      Set objItem = lvTables.ListItems.Add(, , objTable.Identifier, "table", "table")
      objItem.SubItems(1) = Replace(objTable.Comment, vbCrLf, " ")
    End If
  Next objTable
  EndMsg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.GetTables"
End Sub

Private Sub GetColumns()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.GetColumns()", etFullDebug

Dim objItem As ListItem
Dim objColumn As pgColumn

  lvColumns.ListItems.Clear
  StartMsg "Examining Server..."
  For Each objColumn In svr.Databases(lvDatabases.SelectedItem.Text).Tables(lvTables.SelectedItem.Text).Columns
    If Not objColumn.SystemObject Then
      Set objItem = lvColumns.ListItems.Add(, , objColumn.Identifier, "column", "column")
      objItem.SubItems(1) = Replace(objColumn.Comment, vbCrLf, " ")
    End If
  Next objColumn
  EndMsg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.GetColumns"
End Sub

Private Sub GetFormatting()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.GetFormatting()", etFullDebug

  If UCase(RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\Import Wizard", "Trailing Delimiter", "Y")) = "Y" Then
    chkTrailing.Value = 1
  Else
    chkTrailing.Value = 0
  End If
  txtAsciiDelimiter.Text = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\Import Wizard", "Delimiter Character", "44")
  txtAsciiQuote.Text = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\Import Wizard", "Quote Character", "34")
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.GetFormatting"
End Sub


Private Sub GetSubstitutions()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.GetSubstitutions()", etFullDebug

Dim objItem As ListItem
Dim X As Long

  lvSubstitutions.ListItems.Clear
  If Val(RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\" & App.Title & "\Substitution Map", "Entries", "0")) > 0 Then
    For X = 1 To Val(RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\" & App.Title & "\Substitution Map", "Entries", "0"))
      Set objItem = lvSubstitutions.ListItems.Add(, RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\" & App.Title & "\Substitution Map", "Search - " & X, ""), RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\ASCII Exporter\Substitution Map", "Search - " & X, ""))
      objItem.SubItems(1) = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\" & App.Title & "\Substitution Map", "Replace - " & X, "")
    Next
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.GetSubstitutions"
End Sub
  
Private Sub txtDelimiter_Change()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.txtDelimiter_Change()", etFullDebug

  If Not bUpdating Then
    bUpdating = True
    If txtDelimiter.Text = "" Then
      txtAsciiDelimiter.Text = ""
    Else
      txtAsciiDelimiter.Text = Asc(txtDelimiter.Text)
      RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\Import Wizard", "Delimiter Character", regString, txtAsciiDelimiter.Text
    End If
    bUpdating = False
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.txtDelimiter_Change"
End Sub

Private Sub txtQuote_Change()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.txtQuote_Change()", etFullDebug

  If Not bUpdating Then
    bUpdating = True
    If txtQuote.Text = "" Then
      txtAsciiQuote.Text = ""
    Else
      txtAsciiQuote.Text = Asc(txtQuote.Text)
      RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\Import Wizard", "Quote Character", regString, txtAsciiQuote.Text
    End If
    bUpdating = False
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.txtQuote_Change"
End Sub

Private Sub txtAsciiDelimiter_Change()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.txtAsciiDelimiter_Change()", etFullDebug

  If Not bUpdating Then
    bUpdating = True
    If (txtAsciiDelimiter.Text = "") Or (Val(txtAsciiDelimiter.Text) > 255) Or (Val(txtAsciiDelimiter.Text) < 1) Then
      txtDelimiter.Text = ""
    Else
      txtAsciiDelimiter.Text = Val(txtAsciiDelimiter.Text)
      txtAsciiDelimiter.SelStart = Len(txtAsciiDelimiter.Text)
      txtDelimiter.Text = Chr(txtAsciiDelimiter.Text)
      RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\Import Wizard", "Delimiter Character", regString, txtAsciiDelimiter.Text
    End If
    bUpdating = False
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.txtAsciiDelimiter_Change"
End Sub

Private Sub txtAsciiQuote_Change()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.txtAsciiQuote_Change()", etFullDebug

  If Not bUpdating Then
    bUpdating = True
    If (txtAsciiQuote.Text = "") Or (Val(txtAsciiQuote.Text) > 255) Or (Val(txtAsciiQuote.Text) < 1) Then
      txtQuote.Text = ""
    Else
      txtAsciiQuote.Text = Val(txtAsciiQuote.Text)
      txtAsciiQuote.SelStart = Len(txtAsciiQuote.Text)
      txtQuote.Text = Chr(txtAsciiQuote.Text)
      RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\Import Wizard", "Quote Character", regString, txtAsciiQuote.Text
    End If
    bUpdating = False
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.txtAsciiQuote_Change"
End Sub
